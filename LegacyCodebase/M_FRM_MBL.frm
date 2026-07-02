VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form M_FRM_MBL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telecom Interface - PrePaid Telecom Business"
   ClientHeight    =   5610
   ClientLeft      =   1425
   ClientTop       =   1845
   ClientWidth     =   8820
   ForeColor       =   &H00404040&
   Icon            =   "M_FRM_MBL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
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
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
               NumButtonMenus  =   11
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MBL_SUM"
                  Object.Tag             =   "MBL_SUM"
                  Text            =   "&a SUK-SP Summary Report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "INW_REG"
                  Object.Tag             =   "INW_REG"
                  Text            =   "&b Inward Register"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OUT_REG"
                  Object.Tag             =   "OUT_REG"
                  Text            =   "&c Outward Register"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AVT_RPT"
                  Object.Tag             =   "AVT_RPT"
                  Text            =   "&d Activations Report"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NAV_REP"
                  Object.Tag             =   "NAV_REP"
                  Text            =   "&e Non Activations Report"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CAF_REP"
                  Object.Tag             =   "CAF_REP"
                  Text            =   "&f CAF-SEF Receipts Report"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CPN_REP"
                  Object.Tag             =   "CPN_REP"
                  Text            =   "&g CAF-SEF Pending Report"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CHN_REP"
                  Object.Tag             =   "CHN_REP"
                  Text            =   "&h Churn Report"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MBL_STS"
                  Object.Tag             =   "MBL_STS"
                  Text            =   "&i SUK-SP Transaction Record"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MBL_STK"
                  Object.Tag             =   "MBL_STK"
                  Text            =   "&j SUK-SP Stock Report"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "STK_SKU"
                  Object.Tag             =   "STK_SKU"
                  Text            =   "&k Stock Report - SKU"
               EndProperty
            EndProperty
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Add Opening Stock"
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
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Bill"
            Object.ToolTipText     =   "Convert to Bill"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "RS"
            Object.ToolTipText     =   "Return Stocks"
            ImageIndex      =   13
         EndProperty
      EndProperty
      Begin VB.CheckBox Pieces 
         Caption         =   "Check to do in Cases"
         Height          =   195
         Left            =   4800
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   465
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      TabCaption(0)   =   "&a New SUK-SP Numbers Entry"
      TabPicture(0)   =   "M_FRM_MBL.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label40"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label23"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "inward"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LVSRN"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "invdat"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "MS"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtTemp"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtVch"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtSup"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CommonDialog1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "fileimpcmd"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "&b Sales - SUK-SP Numbers"
      TabPicture(1)   =   "M_FRM_MBL.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "M_Prd_Nme"
      Tab(1).Control(1)=   "Total"
      Tab(1).Control(2)=   "Label1(4)"
      Tab(1).Control(3)=   "Label1(5)"
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(5)=   "Label4"
      Tab(1).Control(6)=   "Label1(3)"
      Tab(1).Control(7)=   "Label1(6)"
      Tab(1).Control(8)=   "Label1(10)"
      Tab(1).Control(9)=   "Ptno(0)"
      Tab(1).Control(10)=   "Ptno(1)"
      Tab(1).Control(11)=   "Ptno(2)"
      Tab(1).Control(12)=   "Ptno(3)"
      Tab(1).Control(13)=   "Ptno(4)"
      Tab(1).Control(14)=   "Ptno(5)"
      Tab(1).Control(15)=   "pn"
      Tab(1).Control(16)=   "Label1(7)"
      Tab(1).Control(17)=   "sinvdat"
      Tab(1).Control(18)=   "txtSalInv"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "LV"
      Tab(1).Control(20)=   "cmdsUpdate"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtPartyname"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtSmn"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtSMbl"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "sal_imp"
      Tab(1).Control(25)=   "txtMarket"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "&c Activations Entry"
      TabPicture(2)   =   "M_FRM_MBL.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(9)"
      Tab(2).Control(1)=   "MSA1"
      Tab(2).Control(2)=   "avtdate"
      Tab(2).Control(3)=   "M_FRM_OPT"
      Tab(2).Control(4)=   "Act_imp"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "MSA"
      Tab(2).Control(6)=   "Ynlist"
      Tab(2).Control(7)=   "CMDAUPDATE"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame6"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Vi&ew"
      TabPicture(3)   =   "M_FRM_MBL.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label49"
      Tab(3).Control(1)=   "list4"
      Tab(3).Control(2)=   "list1"
      Tab(3).Control(3)=   "list5"
      Tab(3).Control(4)=   "LIST3"
      Tab(3).Control(5)=   "list2"
      Tab(3).Control(6)=   "LVP"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "&e Numbers Xchange n Change"
      TabPicture(4)   =   "M_FRM_MBL.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label5"
      Tab(4).Control(1)=   "Label6"
      Tab(4).Control(2)=   "txtRetName"
      Tab(4).Control(3)=   "Label7"
      Tab(4).Control(4)=   "Label8"
      Tab(4).Control(5)=   "Label9"
      Tab(4).Control(6)=   "List8"
      Tab(4).Control(7)=   "List7"
      Tab(4).Control(8)=   "List6"
      Tab(4).Control(9)=   "LVR"
      Tab(4).Control(10)=   "LVF"
      Tab(4).Control(11)=   "cmbType"
      Tab(4).Control(12)=   "txtRet"
      Tab(4).Control(13)=   "Command5"
      Tab(4).Control(14)=   "Command6"
      Tab(4).Control(15)=   "MSPC"
      Tab(4).Control(16)=   "Text2"
      Tab(4).ControlCount=   17
      TabCaption(5)   =   "&d CAF - SEF Entry"
      TabPicture(5)   =   "M_FRM_MBL.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "MSC"
      Tab(5).Control(1)=   "CYnlist"
      Tab(5).Control(2)=   "Frame1"
      Tab(5).Control(3)=   "cmdCUpdate"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Frame4"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "&f Churn Entry"
      TabPicture(6)   =   "M_FRM_MBL.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Command4"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "cmdRUpdate"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Frame5"
      Tab(6).Control(3)=   "Frame3"
      Tab(6).Control(4)=   "RYnlist"
      Tab(6).Control(5)=   "MSR"
      Tab(6).Control(6)=   "avtcdate"
      Tab(6).Control(7)=   "Label1(8)"
      Tab(6).ControlCount=   8
      TabCaption(7)   =   "&g Returns"
      TabPicture(7)   =   "M_FRM_MBL.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame7"
      Tab(7).ControlCount=   1
      Begin VB.TextBox txtMarket 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   -70665
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   1710
         Width           =   2070
      End
      Begin VB.CommandButton sal_imp 
         Caption         =   "File i&mport"
         Height          =   345
         Left            =   -67740
         TabIndex        =   141
         Top             =   1245
         Width           =   1305
      End
      Begin VB.Frame Frame7 
         Height          =   4110
         Left            =   -74895
         TabIndex        =   123
         Top             =   795
         Width           =   8430
         Begin VB.Frame Frame8 
            Height          =   420
            Left            =   3030
            TabIndex        =   137
            Top             =   120
            Width           =   2940
            Begin VB.OptionButton Option3 
               Caption         =   "Purchase, Opening"
               Height          =   195
               Left            =   1020
               TabIndex        =   140
               Top             =   150
               Width           =   1680
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Sales"
               Height          =   195
               Left            =   105
               TabIndex        =   138
               Top             =   150
               Value           =   -1  'True
               Width           =   810
            End
         End
         Begin VB.TextBox txtModRet 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1710
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   630
            Width           =   1215
         End
         Begin VB.TextBox txtModInv 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1695
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   1020
            Width           =   1215
         End
         Begin VB.CheckBox chkModDate 
            Caption         =   "Invoice Date"
            Height          =   255
            Left            =   435
            TabIndex        =   125
            Top             =   210
            Width           =   1215
         End
         Begin VB.CommandButton cmdModUpdate 
            Caption         =   "&7 Save"
            Height          =   375
            Left            =   7245
            TabIndex        =   124
            Top             =   585
            Width           =   1035
         End
         Begin MSComctlLib.ListView LVMOD 
            Height          =   2925
            Left            =   3030
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   1035
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   5159
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
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
               Text            =   "Mobile No."
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Activation Status"
               Object.Width           =   3351
            EndProperty
         End
         Begin MSComCtl2.DTPicker modDate 
            Height          =   345
            Left            =   1725
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   210
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   77332483
            CurrentDate     =   37099
         End
         Begin VB.Label TT 
            Caption         =   "Label10"
            Height          =   270
            Left            =   255
            TabIndex        =   150
            Top             =   1050
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label lblno 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   5
            Left            =   795
            TabIndex        =   149
            Top             =   3645
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.Label PRDID 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   5
            Left            =   150
            TabIndex        =   148
            Top             =   3645
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label lblno 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   4
            Left            =   795
            TabIndex        =   147
            Top             =   3240
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.Label PRDID 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   4
            Left            =   150
            TabIndex        =   146
            Top             =   3240
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label PRDID 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   3
            Left            =   165
            TabIndex        =   145
            Top             =   2820
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label PRDID 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   2
            Left            =   165
            TabIndex        =   144
            Top             =   2400
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label PRDID 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   1
            Left            =   150
            TabIndex        =   143
            Top             =   1935
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label PRDID 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   0
            Left            =   165
            TabIndex        =   142
            Top             =   1500
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Select "
            Height          =   195
            Index           =   13
            Left            =   150
            TabIndex        =   139
            Top             =   4095
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Party ID"
            Height          =   195
            Index           =   11
            Left            =   885
            TabIndex        =   136
            Top             =   675
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Voucher ID"
            Height          =   195
            Index           =   12
            Left            =   810
            TabIndex        =   135
            Top             =   1110
            Width           =   810
         End
         Begin VB.Label txtModRetName 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   3030
            TabIndex        =   134
            Top             =   630
            Width           =   2895
         End
         Begin VB.Label lblno 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   0
            Left            =   810
            TabIndex        =   133
            Top             =   1500
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.Label lblno 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   1
            Left            =   795
            TabIndex        =   132
            Top             =   1950
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.Label lblno 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   2
            Left            =   795
            TabIndex        =   131
            Top             =   2400
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.Label lblno 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000C&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            ForeColor       =   &H80000005&
            Height          =   330
            Index           =   3
            Left            =   810
            TabIndex        =   130
            Top             =   2820
            Visible         =   0   'False
            Width           =   2145
         End
      End
      Begin VB.CommandButton fileimpcmd 
         Caption         =   "File &import"
         Height          =   390
         Left            =   7065
         TabIndex        =   121
         Top             =   2580
         Width           =   1350
      End
      Begin VB.Frame Frame6 
         Height          =   495
         Left            =   -74865
         TabIndex        =   111
         Top             =   825
         Width           =   8355
         Begin VB.CheckBox chkAAll 
            Caption         =   "&Z All"
            Height          =   255
            Left            =   150
            TabIndex        =   112
            Top             =   180
            Width           =   975
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "File Imp&ort"
         Height          =   340
         Left            =   -67680
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1215
      End
      Begin VB.CommandButton cmdRUpdate 
         Caption         =   "&6 Save"
         Height          =   340
         Left            =   -69015
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         Height          =   510
         Left            =   -74880
         TabIndex        =   105
         Top             =   675
         Width           =   8400
         Begin VB.CheckBox chkRAll 
            Caption         =   "&X All"
            Height          =   255
            Left            =   150
            TabIndex        =   106
            Top             =   195
            Width           =   975
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   345
         Top             =   2400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame4 
         Height          =   510
         Left            =   -74805
         TabIndex        =   101
         Top             =   810
         Width           =   7065
         Begin VB.CheckBox chkCAll 
            Caption         =   "&Y All"
            Height          =   255
            Left            =   135
            TabIndex        =   102
            Top             =   180
            Width           =   975
         End
      End
      Begin VB.TextBox txtSMbl 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73320
         MaxLength       =   10
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   1710
         Width           =   1320
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73680
         MaxLength       =   10
         TabIndex        =   94
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MSPC 
         Height          =   3255
         Left            =   -74865
         TabIndex        =   93
         Top             =   1785
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   5741
         _Version        =   393216
         BackColor       =   -2147483633
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
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&5 Save "
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67695
         TabIndex        =   89
         Top             =   1365
         Width           =   1290
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Replace With"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67680
         TabIndex        =   88
         Top             =   945
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtRet 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73320
         TabIndex        =   86
         Top             =   1260
         Width           =   1035
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "M_FRM_MBL.frx":0522
         Left            =   -73320
         List            =   "M_FRM_MBL.frx":052C
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   810
         Width           =   2220
      End
      Begin VB.TextBox txtSmn 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   -70665
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   1245
         Width           =   2055
      End
      Begin VB.TextBox txtPartyname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   -70665
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   810
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Height          =   1200
         Left            =   -74880
         TabIndex        =   62
         Top             =   1065
         Width           =   8400
         Begin VB.TextBox txtrSalesman 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1470
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   285
            Width           =   2055
         End
         Begin VB.TextBox txtrRetailer 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1470
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtrInvoice 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5685
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   300
            Width           =   2055
         End
         Begin VB.TextBox txtrMobile 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5670
            MaxLength       =   10
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   690
            Width           =   2055
         End
         Begin VB.OptionButton optrMobile 
            Caption         =   "SUK-SP"
            Height          =   255
            Left            =   4260
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   735
            Width           =   975
         End
         Begin VB.OptionButton optrInvoice 
            Caption         =   "Invoice"
            Height          =   255
            Left            =   4260
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   345
            Width           =   975
         End
         Begin VB.OptionButton optrDate 
            Caption         =   "Date"
            Height          =   255
            Left            =   1890
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1905
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton optrRet 
            Caption         =   "Retailer"
            Height          =   255
            Left            =   90
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   750
            Width           =   975
         End
         Begin VB.OptionButton optrSmn 
            Caption         =   "Salesman"
            Height          =   270
            Left            =   60
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   330
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker ChrDate 
            Height          =   330
            Left            =   3330
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   1860
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   77332483
            CurrentDate     =   37099
         End
      End
      Begin VB.ListBox RYnlist 
         Height          =   255
         ItemData        =   "M_FRM_MBL.frx":055F
         Left            =   -73680
         List            =   "M_FRM_MBL.frx":0569
         TabIndex        =   60
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCUpdate 
         Caption         =   "&4 Save"
         Height          =   340
         Left            =   -67665
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   1230
         Left            =   -74805
         TabIndex        =   52
         Top             =   1215
         Width           =   7065
         Begin VB.TextBox txtcInvoice 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4845
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   255
            Width           =   2055
         End
         Begin VB.OptionButton optInvoice 
            Caption         =   "Invoice"
            Height          =   255
            Left            =   3600
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   300
            Width           =   975
         End
         Begin VB.OptionButton optSmn 
            Caption         =   "Salesman"
            Height          =   255
            Left            =   90
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optRet 
            Caption         =   "Retailer"
            Height          =   255
            Left            =   90
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   735
            Width           =   975
         End
         Begin VB.OptionButton optMobile 
            Caption         =   "SUK-SP"
            Height          =   255
            Left            =   3630
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   735
            Width           =   945
         End
         Begin VB.TextBox txtcMobile 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4860
            MaxLength       =   10
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtcRetailer 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1335
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   705
            Width           =   2055
         End
         Begin VB.TextBox txtcSalesman 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1335
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   285
            Width           =   2055
         End
      End
      Begin VB.ListBox CYnlist 
         Height          =   255
         ItemData        =   "M_FRM_MBL.frx":0576
         Left            =   -73680
         List            =   "M_FRM_MBL.frx":0580
         TabIndex        =   50
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CMDAUPDATE 
         Caption         =   "&3 Save"
         Height          =   340
         Left            =   -68670
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2370
         Width           =   1050
      End
      Begin VB.ListBox Ynlist 
         Height          =   255
         ItemData        =   "M_FRM_MBL.frx":058D
         Left            =   -73680
         List            =   "M_FRM_MBL.frx":0597
         TabIndex        =   48
         Top             =   3000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdsUpdate 
         Caption         =   "&2 Save"
         Height          =   345
         Left            =   -67740
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1305
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&1 Save "
         Height          =   390
         Left            =   7065
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2100
         Width           =   1350
      End
      Begin MSComctlLib.ListView LV 
         Height          =   2910
         Left            =   -73320
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2160
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   5133
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
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
            Text            =   "SUK-SP ID"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SKU Name"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.TextBox txtSalInv 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -73320
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   810
         Width           =   1320
      End
      Begin MSFlexGridLib.MSFlexGrid MSA 
         Height          =   2220
         Left            =   -74865
         TabIndex        =   38
         Top             =   2805
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   3916
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtSup 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1560
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1635
         Width           =   3225
      End
      Begin VB.TextBox txtVch 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1575
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1365
      End
      Begin VB.Frame Frame2 
         Caption         =   "SUK-SP Number"
         Height          =   1335
         Left            =   5955
         TabIndex        =   31
         Top             =   705
         Width           =   2565
         Begin VB.CommandButton cmdGen 
            Caption         =   "&0 Generate"
            Enabled         =   0   'False
            Height          =   390
            Left            =   1080
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   780
            Width           =   1350
         End
         Begin VB.TextBox txtStart 
            Appearance      =   0  'Flat
            Height          =   390
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   33
            Top             =   322
            Width           =   1350
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Start From"
            Height          =   195
            Left            =   135
            TabIndex        =   32
            Top             =   390
            Width           =   720
         End
      End
      Begin VB.TextBox txtTemp 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   2940
         Left            =   1560
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2130
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   5186
         _Version        =   393216
         Cols            =   3
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Act_imp 
         Caption         =   "File Im&port"
         Height          =   340
         Left            =   -67560
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2370
         Width           =   1050
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   -73560
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1140
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "C&reate SUK-SP"
         Height          =   345
         Left            =   1575
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   750
         Width           =   1365
      End
      Begin VB.Frame M_FRM_OPT 
         Height          =   1065
         Left            =   -74865
         TabIndex        =   3
         Top             =   1215
         Width           =   8355
         Begin VB.TextBox txtSalesman 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1275
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtRetailer 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1275
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   615
            Width           =   2055
         End
         Begin VB.TextBox txtInvoice 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4500
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtMobile 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4485
            MaxLength       =   10
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   615
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            Caption         =   "SUK-SP"
            Height          =   255
            Index           =   0
            Left            =   3465
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   660
            Width           =   960
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Invoice"
            Height          =   255
            Index           =   1
            Left            =   3465
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   300
            Width           =   960
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Date"
            Height          =   255
            Index           =   2
            Left            =   4605
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1170
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Retailer"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   660
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Salesman"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   300
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker actDate 
            Height          =   330
            Left            =   6675
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   77332483
            CurrentDate     =   37099
         End
      End
      Begin MSComctlLib.ListView LVP 
         Height          =   3615
         Left            =   -74745
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1155
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComctlLib.ListView list2 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1140
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComctlLib.ListView LIST3 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1140
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Vehicle Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Salesman"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Market"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Route"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView list5 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1140
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
            Text            =   "Challan ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Salesman"
            Object.Width           =   2822
         EndProperty
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1140
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
            Text            =   "Salesman"
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
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Market"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Route"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1140
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product (SKU) Name"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "MRP"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Sale Rate"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Short Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Batch ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Batch Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Pur.Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Mfg. Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Exp. Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker invdat 
         Height          =   330
         Left            =   3450
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   77332483
         CurrentDate     =   37099
      End
      Begin MSComCtl2.DTPicker sinvdat 
         Height          =   330
         Left            =   -73320
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1245
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   77332483
         CurrentDate     =   37099
      End
      Begin MSFlexGridLib.MSFlexGrid MSC 
         Height          =   2415
         Left            =   -74850
         TabIndex        =   51
         Top             =   2580
         Width           =   8430
         _ExtentX        =   14870
         _ExtentY        =   4260
         _Version        =   393216
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSR 
         Height          =   2325
         Left            =   -74880
         TabIndex        =   61
         Top             =   2730
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   4101
         _Version        =   393216
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker avtcdate 
         Height          =   330
         Left            =   -73410
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2325
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   77332483
         CurrentDate     =   37099
      End
      Begin MSComCtl2.DTPicker avtdate 
         Height          =   330
         Left            =   -73320
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   77332483
         CurrentDate     =   37099
      End
      Begin MSComctlLib.ListView LVF 
         Height          =   2895
         Left            =   -74865
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   2160
         Visible         =   0   'False
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComctlLib.ListView LVR 
         Height          =   2895
         Left            =   -70560
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   2160
         Visible         =   0   'False
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.ListBox List6 
         Height          =   255
         Left            =   -73560
         TabIndex        =   90
         Top             =   3240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox List7 
         Height          =   255
         Left            =   -69360
         TabIndex        =   91
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox List8 
         Height          =   450
         Left            =   -68280
         TabIndex        =   92
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ListView LVSRN 
         Height          =   2940
         Left            =   1560
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   2115
         Visible         =   0   'False
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   5186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSA1 
         Height          =   2220
         Left            =   -74880
         TabIndex        =   104
         Top             =   2775
         Visible         =   0   'False
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   3916
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Market"
         Height          =   195
         Index           =   7
         Left            =   -71265
         TabIndex        =   152
         Top             =   1755
         Width           =   495
      End
      Begin VB.Label inward 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inward Nos : 9999"
         ForeColor       =   &H80000005&
         Height          =   330
         Left            =   7065
         TabIndex        =   122
         Top             =   3165
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label pn 
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74835
         TabIndex        =   119
         Top             =   2100
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Ptno 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         ForeColor       =   &H80000005&
         Height          =   330
         Index           =   5
         Left            =   -68535
         TabIndex        =   118
         Top             =   4650
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Ptno 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         ForeColor       =   &H80000005&
         Height          =   330
         Index           =   4
         Left            =   -68535
         TabIndex        =   117
         Top             =   4155
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Ptno 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         ForeColor       =   &H80000005&
         Height          =   330
         Index           =   3
         Left            =   -68535
         TabIndex        =   116
         Top             =   3660
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Ptno 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         ForeColor       =   &H80000005&
         Height          =   330
         Index           =   2
         Left            =   -68535
         TabIndex        =   115
         Top             =   3165
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Ptno 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         ForeColor       =   &H80000005&
         Height          =   330
         Index           =   1
         Left            =   -68550
         TabIndex        =   114
         Top             =   2655
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Ptno 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         ForeColor       =   &H80000005&
         Height          =   330
         Index           =   0
         Left            =   -68550
         TabIndex        =   113
         Top             =   2160
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Press F6 to Shift Between TABS"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -68685
         TabIndex        =   103
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mobile Numbers in Stock"
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
         Left            =   -70500
         TabIndex        =   100
         Top             =   1785
         Width           =   4095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Retailer Mobile Nunbers"
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
         Left            =   -74865
         TabIndex        =   99
         Top             =   1785
         Width           =   4365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Search SUK-SP ID"
         Height          =   195
         Index           =   10
         Left            =   -74790
         TabIndex        =   97
         Top             =   1778
         Width           =   1350
      End
      Begin VB.Label txtRetName 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   -72225
         TabIndex        =   87
         Top             =   1260
         Width           =   2685
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Retailer ID"
         Height          =   195
         Left            =   -74235
         TabIndex        =   85
         Top             =   1305
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Type"
         Height          =   195
         Left            =   -74730
         TabIndex        =   81
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Activation Date"
         Height          =   195
         Index           =   9
         Left            =   -74835
         TabIndex        =   80
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Churn Date"
         Height          =   195
         Index           =   8
         Left            =   -74655
         TabIndex        =   78
         Top             =   2393
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Salesman"
         Height          =   195
         Index           =   6
         Left            =   -71475
         TabIndex        =   76
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Retailer"
         Height          =   195
         Index           =   3
         Left            =   -71325
         TabIndex        =   74
         Top             =   855
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   -73830
         TabIndex        =   47
         Top             =   4870
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sale Quantity:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   46
         Top             =   4870
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Voucher ID"
         Height          =   195
         Index           =   5
         Left            =   -74250
         TabIndex        =   42
         Top             =   878
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Index           =   4
         Left            =   -73785
         TabIndex        =   41
         Top             =   1305
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplied By"
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   34
         Top             =   1725
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Index           =   1
         Left            =   3045
         TabIndex        =   18
         Top             =   1283
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Voucher ID"
         Height          =   210
         Index           =   0
         Left            =   450
         TabIndex        =   17
         Top             =   1282
         Width           =   840
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2370
         TabIndex        =   16
         Top             =   1260
         Width           =   45
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -66525
         TabIndex        =   15
         Top             =   5085
         Width           =   45
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
         TabIndex        =   14
         Top             =   4650
         Width           =   60
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Invoice"
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
         TabIndex        =   13
         Top             =   790
         Width           =   8295
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
         Left            =   -74880
         TabIndex        =   12
         Top             =   5100
         Width           =   45
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   1560
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
            Picture         =   "M_FRM_MBL.frx":05A4
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":0AE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":0EB0
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":1204
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":183C
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":1B90
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":1FF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":22E4
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":26F0
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":27FC
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":2B50
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":2F64
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MBL.frx":34B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4200
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "M_FRM_MBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim RSINV As New ADODB.Recordset
Dim rschk As New ADODB.Recordset
Dim rsAvt As New ADODB.Recordset
Dim RSINL As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
Dim rsact As New ADODB.Recordset
Dim RSSMN As New ADODB.Recordset
Dim rsmar As New ADODB.Recordset
Dim rsimp As New ADODB.Recordset
Dim rsEdt As New ADODB.Recordset
Dim rsEdt1 As New ADODB.Recordset
Dim rsMod As New ADODB.Recordset
Dim rsaMod As New ADODB.Recordset
Dim r As Integer, C As Integer, pqty As Single, TTYP As String, mbid As Boolean
Dim MBLIDY As String, INVIDY As String, DATIDY As String, RETIDY As String, SMNIDY As String
Dim rsdir As New ADODB.Recordset
Dim lst As ListItem
'DECLARE BY PARVATHI
Dim ROWNO As Integer
Dim SUP_IDY() As String
Dim CHK_SUP As Boolean
Dim SRT As String

Private Sub Act_imp_Click()

If MSA.Rows = 1 Then MSA.Rows = 2
Dim RSTEMP As New ADODB.Recordset
Dim RSIMP1 As New ADODB.Recordset

Dim appXL As Object
Dim appX As Object
Dim TYP As String
Dim adate As Date
CommonDialog1.FileName = ""
Set appXL = CreateObject("Excel.Application")
CommonDialog1.FileName = ""
CommonDialog1.filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
CommonDialog1.DialogTitle = "Select Mobile Numbers File"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen

CommonDialog1.CancelError = False
If Not Trim(CommonDialog1.FileName) = "" Then
    If LCase(Right(Trim(CommonDialog1.FileName), 3)) = "xls" Then ' Edited
         appXL.Workbooks.Open (CommonDialog1.FileName)  'GetObject(Trim(CommonDialog1.FileName))
    Else
        MsgBox "Selected File is not a Valid SpreadSheet File"
        Exit Sub
    End If
    
    TYP = appXL.Worksheets(1).Cells(1, 1)
    adate = appXL.Worksheets(1).Cells(2, 1)
    
    If TYP = "Activations" Then
    
    If Val(Mid(adate, 1, 2)) <= 12 And Trim(Mid(adate, 3, 1)) = "/" And Trim(Mid(adate, 6, 1)) = "/" Then
        
        avtdate.Value = Trim(adate)

'        Dim RSTEMP As New ADODB.Recordset
        Dim rstemp1 As New ADODB.Recordset
        Dim rstemp2 As New ADODB.Recordset
        Dim pno1 As Integer
    
        On Error GoTo errh
        Dim mn() As String
        Dim mn1() As String
    
        Dim chk_flg As Boolean
        chk_flg = False
    
        Dim pno As Integer
        Dim row_num As Integer
        
        con.Execute "delete from rp2", qc
        
        If RSTEMP.State = 1 Then RSTEMP.Close
        RSTEMP.Open "select chr_one,chr_two from rp2", con, adOpenKeyset, adLockPessimistic
        
        row_num = 4
        tot_num = 0
        
        pno = 0
        While Len(Trim(appXL.Worksheets(1).Cells(row_num, 1))) > 0
            If Len(Trim(Val(appXL.Worksheets(1).Cells(row_num, 1)))) = 10 Then
            If RSTEMP.State = 1 Then RSTEMP.Close
            RSTEMP.Open "select chr_one,chr_two from rp2 WHERE CHR_ONE='" & Trim(appXL.Worksheets(1).Cells(row_num, 1)) & "'", con, adOpenKeyset, adLockPessimistic
            If RSTEMP.EOF Then
                       
               RSTEMP.AddNew
               RSTEMP(0) = appXL.Worksheets(1).Cells(row_num, 1)
            If Len(Trim(Val(appXL.Worksheets(1).Cells(row_num, 2)))) = 10 Then
               RSTEMP(1) = appXL.Worksheets(1).Cells(row_num, 2)
            End If
               RSTEMP.Update
             tot_num = Val(tot_num) + 1
             End If
             End If
           
            row_num = Val(row_num) + 1
        Wend
        
        For i = 1 To appXL.Workbooks.count
        appXL.Workbooks(i).Close
        Next
        Set appXL = Nothing
        
    Else
        MsgBox "Invalid Date!", vbCritical, "DAS Version 4.0"
        For i = 1 To appXL.Workbooks.count
        appXL.Workbooks(i).Close
        Next
        Set appXL = Nothing
        Exit Sub
    End If
Else
    MsgBox "Invalid Date!", vbCritical, "DAS Version 4.0"
    For i = 1 To appXL.Workbooks.count
    appXL.Workbooks(i).Close
    Next
    Set appXL = Nothing
    Exit Sub
End If

        'Label1(9).Caption = Label1(9).Caption & ":"
        If RSINL.State = 1 Then RSINL.Close
        Set RSINL = Nothing
        'date wise also
        RSINL.Open "select mbl_idy,sal_idy,inv_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where MBL_IDY IN (SELECT CHR_ONE FROM RP2) AND sal_typ='SAL' AND (avt_sts is null or avt_sts='No' or avt_sts='') order by mbl_idy,sal_dat", con, adOpenKeyset, adLockPessimistic
        If RSINL.RecordCount > 0 Then
            MSA.FixedRows = 0
            MSA.Rows = 1
            MSA.Cols = 8
            MSA.TextMatrix(0, 0) = "SNO"
            MSA.ColWidth(0) = 700
            MSA.TextMatrix(0, 1) = "Mobile ID"
            MSA.ColWidth(1) = 1500
            MSA.ColAlignment(1) = 1
            MSA.TextMatrix(0, 2) = "Invoice ID"
            MSA.ColWidth(2) = 1600
            MSA.TextMatrix(0, 3) = "Status"
            MSA.ColWidth(3) = 800
            MSA.TextMatrix(0, 4) = "SKU Name"
            MSA.ColWidth(4) = 2000
            MSA.TextMatrix(0, 5) = "Party"
            MSA.ColWidth(5) = 2000
            MSA.TextMatrix(0, 6) = "Invoice Date"
            MSA.ColWidth(6) = 1200
            MSA.TextMatrix(0, 7) = "New Number"
            MSA.ColWidth(6) = 1200
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
            MSA.Rows = RSINL.RecordCount + 1
            r = 1
            While Not RSINL.EOF
                MSA.TextMatrix(r, 0) = r
                MSA.TextMatrix(r, 1) = Trim(RSINL(0))
                If IsNull(RSINL(1)) = True Then
                MSA.TextMatrix(r, 2) = " "
                Else
                MSA.TextMatrix(r, 2) = Trim(RSINL(1))
                End If
                MSA.TextMatrix(r, 3) = "Yes" 'Trim(RSINL(6)) & ""
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsprd.EOF Then MSA.TextMatrix(r, 4) = Trim(rsprd(0))
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act where ACT_IDY='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsact.EOF Then MSA.TextMatrix(r, 5) = Trim(rsact(1))
                MSA.TextMatrix(r, 6) = Trim(RSINL(2))
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select chr_two from rp2 where chr_one='" & Trim(RSINL(0)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsact.EOF Then
                If IsNull(rsact(0)) = False Then MSA.TextMatrix(r, 7) = Trim(rsact(0))
                End If
                r = r + 1
                RSINL.MoveNext
            Wend
            CMDAUPDATE.SetFocus
End If
con.Execute "delete from rp2", qc
Exit Sub

Else
    MsgBox "File Selection Cancelled!", vbInformation, "DAS Version 4.0"
    Exit Sub
End If

errh:
MsgBox "Select a Valid File!", vbInformation, "DAS Version 4.0"


End Sub

Private Sub actDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    actDate = Date$
    txtSalesman.SetFocus
    Option1(4).Value = True
End If
If KeyCode = 13 Then
    Label49 = "Select Invoice (Date)"
    If rsAvt.State = 1 Then rsAvt.Close
    Set rsAvt = Nothing
    rsAvt.Open "select mbl_idy,sal_idy,sal_dat,cst_idy,prd_idy,txn_typ,avt_sts from avt where sal_typ like 'SAL' and (avt_sts is null or avt_sts='' or avt_sts <> 'Yes') and sal_dat=ctod('" & actDate & "') order by sal_idy", con, adOpenKeyset, adLockPessimistic
    If rsAvt.RecordCount > 0 Then
        MSA.Rows = 1
        MSA.Cols = 6
        MSA.TextMatrix(0, 0) = "Mobile ID"
        MSA.ColWidth(0) = 1300
        MSA.ColAlignment(0) = 1
        MSA.TextMatrix(0, 1) = "Invoice ID"
        MSA.ColWidth(1) = 1300
        MSA.TextMatrix(0, 2) = "Status"
        MSA.ColWidth(2) = 800
        MSA.TextMatrix(0, 3) = "SKU Name"
        MSA.ColWidth(3) = 2200
        MSA.TextMatrix(0, 4) = "Party"
        MSA.ColWidth(4) = 1740
        MSA.TextMatrix(0, 5) = "Invoice Date"
        MSA.ColWidth(5) = 1200
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
        MSA.Rows = rsAvt.RecordCount + 1
'                MSA.FixedRows = 1
        r = 1
        While Not rsAvt.EOF
            
            MSA.TextMatrix(r, 0) = Trim(rsAvt(0))
            MSA.TextMatrix(r, 1) = Trim(rsAvt(1))
            MSA.TextMatrix(r, 2) = "No" 'Trim(rsavt(6)) & ""
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsprd.EOF Then MSA.TextMatrix(r, 3) = Trim(rsprd(0))
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act where act_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then MSA.TextMatrix(r, 4) = Trim(rsact(1))
            MSA.TextMatrix(r, 5) = Trim(rsAvt(2))
            r = r + 1
            rsAvt.MoveNext
        Wend
        SSTab1.Tab = 2
        LVP.Visible = False
        SSTab1.TabEnabled(3) = False
        MSA.Col = 2
        MSA.SetFocus
    Else
        MsgBox "No Invoices Found", vbInformation + vbOKOnly, "DAS Version 4.0"
        SSTab1.Tab = 2
        'actDate.SetFocus
        'Option1(2).Value = True
        chkAAll.SetFocus
        chkAAll.Value = 1
        SSTab1.TabEnabled(3) = False
    End If
End If
End Sub

Private Sub avtcdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MSR.Rows > 1 Then
        MSR.Col = 3
        MSR.SetFocus
    End If
End If
End Sub

Private Sub avtcdate_LostFocus()
If SSTab1.Tab <> 6 Then
            MSR.Rows = 1
            chkRAll.Enabled = True
            chkRAll.SetFocus
            chkRAll.Value = 1
            txtrSalesman.Text = ""
            txtrInvoice.Text = ""
            txtrRetailer.Text = ""
            txtrMobile.Text = ""
    If SSTab1.Tab = 5 Then
    If chkCAll.Enabled = False Then chkCAll.Enabled = True
    chkCAll.SetFocus
    ElseIf SSTab1.Tab = 2 Then
    If chkAAll.Enabled = False Then chkAAll.Enabled = True
    chkAAll.SetFocus
    End If
            End If
End Sub

Private Sub avtdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MSA.Rows > 1 Then
        MSA.Col = 3
        MSA.SetFocus
    End If
End If
End Sub

Private Sub cafdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    cafdate = Date$
    optSmn.Value = True
    txtcSalesman.SetFocus
End If
If KeyCode = 13 Then
'    Dim ls As ListItem
    Label49 = "Select Invoice (Date) for CAF"
'    SSTab1.TabEnabled(3) = True
'    If RSINL.State = 1 Then RSINL.Close
'    RSINL.Open "select sal_idy,sal_dat,cst_idy,prd_idy,txn_typ,avt_sts from avt where sal_typ like 'SAL' and avt_sts='Yes' and (caf_sts is null or caf_sts='' or caf_sts<>'Yes') and sal_dat=ctod('" & actDate & "') group by sal_idy", CON, adOpenKeyset, adlockpessimistic
'    If RSINL.RecordCount > 0 Then
'        LVP.ColumnHeaders.Clear
'        LVP.ColumnHeaders.Add , , "Invoice ID"
'        LVP.ColumnHeaders.Add , , "SKU Name", 2500
'        LVP.ColumnHeaders.Add , , "Party", 2500
'        LVP.ColumnHeaders.Add , , "Invoice Date", 1200
'        LVP.ListItems.Clear
'        If rsprd.State = 1 Then rsprd.Close
'        rsprd.Open "select prd_nme,prd_idy from prd", CON, adOpenKeyset, adlockpessimistic
'        If rsact.State = 1 Then rsact.Close
'        rsact.Open "select act_idy,act_nme from act", CON, adOpenKeyset, adlockpessimistic
'        While Not RSINL.EOF
'            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
'            If rsact.RecordCount > 0 Then rsact.MoveFirst
'            Set ls = LVP.ListItems.Add(, , RSINL(0))
'            rsprd.Find "prd_idy='" & Trim(RSINL(3)) & "'"
'            If Not rsprd.EOF Then ls.ListSubItems.Add , , Trim(rsprd(0))
'            rsact.Find "act_idy='" & Trim(RSINL(3)) & "'"
'            If Not rsact.EOF Then ls.ListSubItems.Add , , rsact(1)
'            ls.ListSubItems.Add , , RSINL(1)
'            RSINL.MoveNext
'        Wend
'        SSTab1.TabEnabled(3) = True
'        LVP.Visible = True
'        LVP.SetFocus
'        SSTab1.Tab = 3
'    Else
'        MsgBox "No Invoices Found", vbInformation + vbOKOnly, "DAS Version 4.0"
'        SSTab1.Tab = 5
'        cafdate.SetFocus
'        SSTab1.TabEnabled(3) = False
'    End If
    If rsAvt.State = 1 Then rsAvt.Close
    Set rsAvt = Nothing
    rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' and avt_sts='Yes' and (caf_sts is null or caf_sts<>'Yes' or caf_sts='') and avt_dat=ctod('" & cafdate & "') order by mbl_idy", con, adOpenKeyset, adLockPessimistic
    If rsAvt.RecordCount > 0 Then
        MSC.Rows = 1
        MSC.Cols = 6
        MSC.TextMatrix(0, 0) = "Mobile ID"
        MSC.ColWidth(0) = 1300
        MSC.ColAlignment(0) = 1
        MSC.TextMatrix(0, 1) = "Invoice ID"
        MSC.ColWidth(1) = 1300
        MSC.TextMatrix(0, 2) = "Status"
        MSC.ColWidth(2) = 800
        MSC.TextMatrix(0, 3) = "SKU Name"
        MSC.ColWidth(3) = 2200
        MSC.TextMatrix(0, 4) = "Party"
        MSC.ColWidth(4) = 1740
        MSC.TextMatrix(0, 5) = "Act. Date"
        MSC.ColWidth(5) = 1200
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
        MSC.Rows = rsAvt.RecordCount + 1
'                MSA.FixedRows = 1
        r = 1
        While Not rsAvt.EOF
            MSC.TextMatrix(r, 0) = Trim(rsAvt(0))
            MSC.TextMatrix(r, 1) = Trim(rsAvt(1))
            MSC.TextMatrix(r, 2) = "No" 'Trim(rsavt(6)) & ""
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsprd.EOF Then MSC.TextMatrix(r, 3) = Trim(rsprd(0))
            rsact.Open "select act_idy,act_nme from act where act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then MSC.TextMatrix(r, 4) = Trim(rsact(1))
            MSC.TextMatrix(r, 5) = Trim(rsAvt(2))
            r = r + 1
            rsAvt.MoveNext
        Wend
        SSTab1.Tab = 5
        LVP.Visible = False
        SSTab1.TabEnabled(3) = False
        MSC.Col = 2
        MSC.SetFocus
    Else
        MsgBox "No Invoice Found in the selected Date!", vbInformation + vbOKOnly, "DAS Version 4.0"
        cafdate.SetFocus
        Exit Sub
    End If
End If
End Sub



Private Sub avtdate_KeyPress(K As Integer)
If K = 27 Then
If SSTab1.Tab <> 2 Then SSTab1.Tab = 2
If chkAAll.Enabled = False Then chkAAll.Enabled = True
chkAAll.SetFocus
chkAAll.Value = 1
MSA.Rows = 1
Option1(0).Enabled = True
Option1(1).Enabled = True
Option1(3).Enabled = True
Option1(4).Enabled = True
txtSalesman = ""
txtRetailer = ""
txtInvoice = ""
txtMobile = ""
End If

End Sub

Private Sub avtdate_LostFocus()
'TOTAL CODE ADDED BY PPKB
If SSTab1.Tab <> 2 Then
        MSA.Rows = 1
        txtMobile = ""
        txtInvoice = ""
        txtRetailer = ""
        txtSalesman = ""
        End If
End Sub

Private Sub chkAAll_Click()
'code added by pkb
MSA.CLEAR
If chkAAll.Value = 1 Then
txtSalesman = ""
txtInvoice = ""
txtRetailer = ""
txtMobile = ""
Option1(0).Value = 0
Option1(1).Value = 0
Option1(3).Value = 0
Option1(4).Value = 0
End If
End Sub

Private Sub chkAAll_GotFocus()
If chkAAll.Value = 1 Then
Option1(0).Value = 0
Option1(1).Value = 0
Option1(3).Value = 0
Option1(4).Value = 0
End If
End Sub

Private Sub chkAAll_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If chkAAll.Value = 1 Then
        'disable
        M_FRM_OPT.Enabled = False
        Option1(0).Enabled = False
        Option1(1).Enabled = False
        Option1(3).Enabled = False
        Option1(4).Enabled = False
        txtSalesman.Visible = False
        txtRetailer.Visible = False
        txtInvoice.Visible = False
        txtMobile.Visible = False
        txtSalesman.Enabled = False
        txtRetailer.Enabled = False
        txtInvoice.Enabled = False
        txtMobile.Enabled = False
        'disable
        
        If RSINL.State = 1 Then RSINL.Close
        'Set RSINL = Nothing
        'date wise also
        RSINL.Open "select mbl_idy,sal_idy,inv_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where LEN(MBL_IDY) > 0 AND AVT_STS IS NULL order by mbl_idy,sal_dat", con, adOpenKeyset, adLockPessimistic
        If (RSINL.RecordCount > 0) Then
            MSA.FixedRows = 0
            MSA.Rows = 1
            MSA.Cols = 7
            MSA.TextMatrix(0, 0) = "SNO"
            MSA.ColWidth(0) = 700
            MSA.TextMatrix(0, 1) = "Mobile ID"
            MSA.ColWidth(1) = 1500
            MSA.ColAlignment(1) = 1
            MSA.TextMatrix(0, 2) = "Invoice ID"
            MSA.ColWidth(2) = 1600
            MSA.TextMatrix(0, 3) = "Status"
            MSA.ColWidth(3) = 800
            MSA.TextMatrix(0, 4) = "SKU Name"
            MSA.ColWidth(4) = 2000
            MSA.TextMatrix(0, 5) = "Party"
            MSA.ColWidth(5) = 2000
            MSA.TextMatrix(0, 6) = "Invoice Date"
            MSA.ColWidth(6) = 1200
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
            MSA.Rows = RSINL.RecordCount + 1
            MSA.FixedRows = 1
            r = 1
            While Not RSINL.EOF
                MSA.TextMatrix(r, 0) = r
                MSA.TextMatrix(r, 1) = Trim(RSINL(0))
                If IsNull(RSINL(1)) = True Then
                MSA.TextMatrix(r, 2) = " "
                Else
                MSA.TextMatrix(r, 2) = Trim(RSINL(1))
                End If
                MSA.TextMatrix(r, 3) = "NO" 'Trim(RSINL(6)) & ""
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsprd.EOF Then MSA.TextMatrix(r, 4) = Trim(rsprd(0))
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act where ACT_IDY='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsact.EOF Then MSA.TextMatrix(r, 5) = Trim(rsact(1))
                MSA.TextMatrix(r, 6) = Trim(RSINL(2))
                r = r + 1
                RSINL.MoveNext
            Wend
            avtdate.SetFocus

        Else
            MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
            If chkAAll.Enabled = False Then chkAAll.Enabled = True
            chkAAll.SetFocus
            chkAAll.Value = 1
        End If
    Else
        'enable
        chkAAll.Enabled = False
        M_FRM_OPT.Enabled = True
        Option1(0).Enabled = True
        Option1(1).Enabled = True
        Option1(3).Enabled = True
        Option1(4).Enabled = True
        txtSalesman.Enabled = True
        txtRetailer.Enabled = True
        txtInvoice.Enabled = True
        txtMobile.Enabled = True
        txtSalesman.Visible = True
        txtRetailer.Visible = True
        txtInvoice.Visible = True
        txtMobile.Visible = True
        'enable
        Option1(4).Value = True
        txtSalesman.SetFocus
    End If
End If
End Sub

Private Sub chkAAll_LostFocus()
If chkAAll.Value = 1 Then
txtSalesman.Enabled = False
txtInvoice.Enabled = False
txtRetailer.Enabled = False
txtMobile.Enabled = False
Option1(0).Enabled = False
Option1(1).Enabled = False
Option1(3).Enabled = False
Option1(4).Enabled = False
If chkAAll.Enabled = False Then chkAAll.Enabled = True
End If
End Sub

Private Sub chkCAll_Click()
'code added by pkb
MSC.CLEAR
If chkCAll.Value = 1 Then
txtcSalesman = ""
txtcInvoice = ""
txtcRetailer = ""
txtcMobile = ""
optSmn.Value = False
optInvoice.Value = False
optRet.Value = False
optMobile.Value = False
End If
End Sub

'Private Sub chkAAll_LostFocus()
'If chkAAll.Value = 1 Then
'
'        'disable
'        M_FRM_OPT.Enabled = False
'        Option1(0).Enabled = False
'        Option1(1).Enabled = False
'        Option1(3).Enabled = False
'        Option1(4).Enabled = False
'        txtSalesman.Visible = False
'        txtRetailer.Visible = False
'        txtInvoice.Visible = False
'        txtMobile.Visible = False
'        txtSalesman.Enabled = False
'        txtRetailer.Enabled = False
'        txtInvoice.Enabled = False
'        txtMobile.Enabled = False
'    Else
'        chkAAll.Enabled = False
'        M_FRM_OPT.Enabled = True
'        Option1(0).Enabled = True
'        Option1(1).Enabled = True
'        Option1(3).Enabled = True
'        Option1(4).Enabled = True
'        txtSalesman.Enabled = True
'        txtRetailer.Enabled = True
'        txtInvoice.Enabled = True
'        txtMobile.Enabled = True
'        txtSalesman.Visible = True
'        txtRetailer.Visible = True
'        txtInvoice.Visible = True
'        txtMobile.Visible = True
'        'enable
'        Option1(4).Value = True
'        txtSalesman.SetFocus
'        End If
'End Sub

Private Sub chkCAll_GotFocus()
If chkCAll.Value = 1 Then
optSmn.Value = False
optInvoice.Value = False
optRet.Value = False
optMobile.Value = False
End If
End Sub

Private Sub chkCAll_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If chkCAll.Value = 1 Then
        'modified by Parvathi
        If Frame1.Enabled = False Then Frame1.Enabled = True
            optSmn.Value = 0
            optInvoice.Value = 0
            optRet.Value = 0
            optMobile.Value = 0
            txtcInvoice.Text = ""
            txtcSalesman.Text = ""
            txtcRetailer.Text = ""
            txtcMobile.Text = ""
            
            txtcInvoice.Visible = False
            txtcSalesman.Visible = False
            txtcRetailer.Visible = False
            txtcMobile.Visible = False
            
                        
            optSmn.Enabled = False
            optInvoice.Enabled = False
            optRet.Enabled = False
            optMobile.Enabled = False
         
        If RSINL.State = 1 Then RSINL.Close
        Set RSINL = Nothing
        RSINL.Open "select mbl_idy,sal_idy,inv_dat,prd_idy,cst_idy,txn_typ,avt_sts,caf_sts,AVT_DAT from avt where sal_typ='SAL' and avt_sts='Yes' and (caf_sts is null or caf_sts<>'Yes' or caf_sts='')", con, adOpenKeyset, adLockPessimistic
        'RSINL.Open "select mbl_idy,sal_idy,inv_dat,prd_idy,cst_idy,txn_typ,avt_sts,caf_sts from avt where sal_typ='SAL' AND avt_sts='Yes' and (caf_sts is null or caf_sts<>'Yes' or caf_sts='')", CON, adOpenKeyset, adlockpessimistic
        If RSINL.RecordCount > 0 Then
            MSC.Rows = 1
            MSC.Cols = 7
            MSC.TextMatrix(0, 0) = "SNO"
            MSC.ColWidth(0) = 700
            MSC.TextMatrix(0, 1) = "Mobile ID"
            MSC.ColWidth(1) = 1300
            MSC.ColAlignment(1) = 1
            MSC.TextMatrix(0, 2) = "Invoice ID"
            MSC.ColWidth(2) = 1300
            MSC.TextMatrix(0, 3) = "Status"
            MSC.ColWidth(3) = 800
            MSC.TextMatrix(0, 4) = "SKU Name"
            MSC.ColWidth(4) = 2200
            MSC.TextMatrix(0, 5) = "Party"
            MSC.ColWidth(5) = 1740
            MSC.TextMatrix(0, 6) = "Act. Date"
            MSC.ColWidth(6) = 1200
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
            MSC.Rows = RSINL.RecordCount + 1
    '        MSA.FixedRows = 1
            r = 1
            While Not RSINL.EOF
                MSC.TextMatrix(r, 0) = r
                MSC.TextMatrix(r, 1) = Trim(RSINL(0))
                If IsNull(RSINL(1)) = True Then
                 MSC.TextMatrix(r, 2) = " "
                Else
                MSC.TextMatrix(r, 2) = Trim(RSINL(1))
                End If
                MSC.TextMatrix(r, 3) = "No" 'Trim(RSINL(7)) & ""
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsprd.EOF Then MSC.TextMatrix(r, 4) = Trim(rsprd(0))
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act where act_idy='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsact.EOF Then MSC.TextMatrix(r, 5) = Trim(rsact(1))
                MSC.TextMatrix(r, 6) = Trim(RSINL(8))
                r = r + 1
                RSINL.MoveNext
            Wend
            Frame1.Enabled = False
            MSC.Col = 3
            MSC.SetFocus
        Else
        
            MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
            SSTab1.Tab = 5
            'cafdate.SetFocus
            'SSTab1.TabEnabled(3) = False
            chkCAll.SetFocus
            chkCAll.Value = 1
        End If
    Else
        SSTab1.Tab = 5
        Frame1.Enabled = True
        If txtcSalesman.Visible = False Then txtcSalesman.Visible = True
        If txtcSalesman.Enabled = False Then txtcSalesman.Enabled = True
        txtcSalesman.SetFocus
        chkCAll.Enabled = False
        optSmn.Value = True
        'Modified by Parvathi
        txtcInvoice.Visible = True
        txtcSalesman.Visible = True
        txtcRetailer.Visible = True
        txtcMobile.Visible = True
        
        optSmn.Enabled = True
        optInvoice.Enabled = True
        optRet.Enabled = True
        optMobile.Enabled = True
        
        MSC.Rows = 1
        MSC.Cols = 6
        MSC.TextMatrix(0, 0) = "Mobile ID"
        MSC.ColWidth(0) = 1300
        MSC.ColAlignment(0) = 1
        MSC.TextMatrix(0, 1) = "Invoice ID"
        MSC.ColWidth(1) = 1300
        MSC.TextMatrix(0, 2) = "Status"
        MSC.ColWidth(2) = 800
        MSC.TextMatrix(0, 3) = "SKU Name"
        MSC.ColWidth(3) = 2200
        MSC.TextMatrix(0, 4) = "Party"
        MSC.ColWidth(4) = 1740
        MSC.TextMatrix(0, 5) = "Act. Date"
        MSC.ColWidth(5) = 1200
            
        'modified by pKB
        txtcSalesman.SetFocus
    End If
End If
End Sub

Private Sub chkCAll_KeyPress(K As Integer)
'If K = 27 Then
'SSTab1.Tab = 2
'End If
End Sub

Private Sub chkCAll_LostFocus()
If chkCAll.Value = 1 Then
txtcSalesman.Enabled = False
txtcInvoice.Enabled = False
txtcRetailer.Enabled = False
txtcMobile.Enabled = False
optSmn.Enabled = False
        optInvoice.Enabled = False
        optRet.Enabled = False
        optMobile.Enabled = False
End If
End Sub

'Private Sub chkCAll_LostFocus()
'If chkCAll.Value = 1 Then
'       ' modified by PKB
'        If Frame1.Enabled = False Then Frame1.Enabled = True
'            optSmn.Value = 0
'            optInvoice.Value = 0
'            optRet.Value = 0
'            optMobile.Value = 0
'            txtcInvoice.Text = ""
'            txtcSalesman.Text = ""
'            txtcRetailer.Text = ""
'            txtcMobile.Text = ""
'
'            txtcInvoice.Visible = False
'            txtcSalesman.Visible = False
'            txtcRetailer.Visible = False
'            txtcMobile.Visible = False
'
'
'            optSmn.Enabled = False
'            optInvoice.Enabled = False
'            optRet.Enabled = False
'            optMobile.Enabled = False
'Else
'    chkCAll.Enabled = False
'        Frame1.Enabled = True
'        optSmn.Value = True
'        'Modified by PKB
'        txtcInvoice.Visible = True
'        txtcSalesman.Visible = True
'        txtcRetailer.Visible = True
'        txtcMobile.Visible = True
'
'        optSmn.Enabled = True
'        optInvoice.Enabled = True
'        optRet.Enabled = True
'        optMobile.Enabled = True
'
'        MSC.Rows = 1
'        MSC.Cols = 6
'        MSC.TextMatrix(0, 0) = "Mobile ID"
'        MSC.ColWidth(0) = 1300
'        MSC.ColAlignment(0) = 1
'        MSC.TextMatrix(0, 1) = "Invoice ID"
'        MSC.ColWidth(1) = 1300
'        MSC.TextMatrix(0, 2) = "Status"
'        MSC.ColWidth(2) = 800
'        MSC.TextMatrix(0, 3) = "SKU Name"
'        MSC.ColWidth(3) = 2200
'        MSC.TextMatrix(0, 4) = "Party"
'        MSC.ColWidth(4) = 1740
'        MSC.TextMatrix(0, 5) = "Act. Date"
'        MSC.ColWidth(5) = 1200
'
'        'modified by PKB
'        txtcSalesman.SetFocus
'    End If
'End Sub

'Private Sub chkModDate_GotFocus()
'    lblno(0).Visible = False
'    lblno(3).Visible = False
'    lblno(2).Visible = False
'    lblno(1).Visible = False
'End Sub

Private Sub chkModDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If chkModDate.Value = 1 Then
        modDate.Enabled = True
        modDate.SetFocus
    Else
        Option2.SetFocus
'        txtModRet.Enabled = True
'        txtModRet.SetFocus
    End If
End If
End Sub

Private Sub chkRAll_Click()
MSR.CLEAR
'code added by pkb
If chkRAll.Value = 1 Then
optrSmn.Value = False
optrRet.Value = False
optrInvoice.Value = False
optrMobile.Value = False
txtrSalesman = ""
txtrInvoice = ""
txtrRetailer = ""
txtrMobile = ""
End If
End Sub

Private Sub chkRAll_GotFocus()
If chkRAll.Value = 1 Then
optrSmn.Value = False
optrRet.Value = False
optrInvoice.Value = False
optrMobile.Value = False
End If
End Sub

Private Sub chkRAll_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        If chkRAll.Value = 1 Then
        Frame3.Enabled = False
        'Disable
        optrSmn.Enabled = False
        optrRet.Enabled = False
        optrInvoice.Enabled = False
        optrMobile.Enabled = False
        txtrSalesman.Enabled = False
        txtrSalesman.Visible = False
        txtrRetailer.Visible = False
        txtrInvoice.Visible = False
        txtrMobile.Visible = False
        txtrRetailer.Enabled = False
        txtrInvoice.Enabled = False
        txtrMobile.Enabled = False
        'Disable
        If RSINL.State = 1 Then RSINL.Close
        RSINL.Open "select mbl_idy,sal_idy,inv_dat,prd_idy,cst_idy,txn_typ,avt_sts,caf_sts from avt where sal_typ='SAL' AND avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='')", con, adOpenKeyset, adLockPessimistic
        If RSINL.RecordCount > 0 Then
            MSR.Rows = 1
            MSR.Cols = 7
            MSR.TextMatrix(0, 0) = "SNO"
            MSR.ColWidth(0) = 700
            MSR.TextMatrix(0, 1) = "Mobile ID"
            MSR.ColWidth(1) = 1500
            MSR.ColAlignment(1) = 1
            MSR.TextMatrix(0, 2) = "Invoice ID"
            MSR.ColWidth(2) = 1600
            MSR.TextMatrix(0, 3) = "Status"
            MSR.ColWidth(3) = 800
            MSR.TextMatrix(0, 4) = "SKU Name"
            MSR.ColWidth(4) = 2000
            MSR.TextMatrix(0, 5) = "Party"
            MSR.ColWidth(5) = 2000
            MSR.TextMatrix(0, 6) = "Invoice Date"
            MSR.ColWidth(6) = 1200
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
            MSR.Rows = RSINL.RecordCount + 1
    '        MSA.FixedRows = 1
            r = 1
            While Not RSINL.EOF
                MSR.TextMatrix(r, 0) = r
                MSR.TextMatrix(r, 1) = Trim(RSINL(0))
                If IsNull(RSINL(1)) = True Then
                MSR.TextMatrix(r, 2) = " "
                Else
                MSR.TextMatrix(r, 2) = Trim(RSINL(1))
                End If
                MSR.TextMatrix(r, 3) = "No" 'Trim(RSINL(6)) & ""
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsprd.EOF Then MSR.TextMatrix(r, 4) = Trim(rsprd(0))
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act where act_idy='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsact.EOF Then MSR.TextMatrix(r, 5) = Trim(rsact(1))
                MSR.TextMatrix(r, 6) = Trim(RSINL(2))
                r = r + 1
                RSINL.MoveNext
            Wend
            'MSR.Col = 2
            'MSR.SetFocus
            avtcdate.SetFocus
        Else
            MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
            SSTab1.Tab = 6
            chkRAll.SetFocus
            SSTab1.TabEnabled(3) = False
        End If
    Else
        SSTab1.Tab = 6
        Frame3.Enabled = True
        If txtrSalesman.Visible = False Then txtrSalesman.Visible = True
        If txtrSalesman.Enabled = False Then txtrSalesman.Enabled = True
        txtrSalesman.SetFocus
        chkRAll.Enabled = False
        'enable
        optrSmn.Enabled = True
        optrRet.Enabled = True
        optrInvoice.Enabled = True
        optrMobile.Enabled = True
        txtrSalesman.Enabled = True
        txtrSalesman.Visible = True
        txtrRetailer.Visible = True
        txtrInvoice.Visible = True
        txtrMobile.Visible = True
        txtrRetailer.Enabled = True
        txtrInvoice.Enabled = True
        txtrMobile.Enabled = True
        'enable
        optrSmn.Value = True
        txtrSalesman.SetFocus   ' TO BE CHECK BY PNR
    End If
End If
End Sub

Private Sub chkRAll_LostFocus()
If chkRAll.Value = 1 Then
optrSmn.Value = False
optrRet.Value = False
optrInvoice.Value = False
optrMobile.Value = False
End If
End Sub

'MODIFIED BY PKB
'Private Sub chkRAll_LostFocus()
'        If chkRAll.Value = 1 Then
'        Frame3.Enabled = False
'        'Disable
'        optrSmn.Enabled = False
'        optrRet.Enabled = False
'        optrInvoice.Enabled = False
'        optrMobile.Enabled = False
'        txtrSalesman.Enabled = False
'        txtrSalesman.Visible = False
'        txtrRetailer.Visible = False
'        txtrInvoice.Visible = False
'        txtrMobile.Visible = False
'        txtrRetailer.Enabled = False
'        txtrInvoice.Enabled = False
'        txtrMobile.Enabled = False
'
''
'Else
'        chkRAll.Enabled = False
'        Frame3.Enabled = True
'        'enable
'        optrSmn.Enabled = True
'        optrRet.Enabled = True
'        optrInvoice.Enabled = True
'        optrMobile.Enabled = True
'        txtrSalesman.Enabled = True
'        txtrSalesman.Visible = True
'        txtrRetailer.Visible = True
'        txtrInvoice.Visible = True
'        txtrMobile.Visible = True
'        txtrRetailer.Enabled = True
'        txtrInvoice.Enabled = True
'        txtrMobile.Enabled = True
'        'enable
'        optrSmn.Value = True
'        txtrSalesman.SetFocus
'End If
'End Sub

Private Sub ChrDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    ChrDate = Date$
    optrSmn.Value = True
    txtrSalesman.SetFocus
End If
If KeyCode = 13 Then
    Label49 = "Select Invoice (Date) for Churn"
'    SSTab1.TabEnabled(3) = True
'    If RSINL.State = 1 Then RSINL.Close
'    RSINL.Open "select sal_idy,sal_dat,cst_idy,prd_idy,txn_typ,avt_sts from avt where sal_typ like 'SAL' and avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') and sal_dat=ctod('" & actDate & "') group by sal_idy", CON, adOpenKeyset, adlockpessimistic
'    If RSINL.RecordCount > 0 Then
'        LVP.ColumnHeaders.Clear
'        LVP.ColumnHeaders.Add , , "Invoice ID"
'        LVP.ColumnHeaders.Add , , "SKU Name", 2500
'        LVP.ColumnHeaders.Add , , "Party", 2500
'        LVP.ColumnHeaders.Add , , "Invoice Date", 1200
'        LVP.ListItems.Clear
'        If rsprd.State = 1 Then rsprd.Close
'        rsprd.Open "select prd_nme,prd_idy from prd", CON, adOpenKeyset, adlockpessimistic
'        If rsact.State = 1 Then rsact.Close
'        rsact.Open "select act_idy,act_nme from act", CON, adOpenKeyset, adlockpessimistic
'        While Not RSINL.EOF
'            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
'            If rsact.RecordCount > 0 Then rsact.MoveFirst
'            Set ls = LVP.ListItems.Add(, , RSINL(0))
'            rsprd.Find "prd_idy='" & Trim(RSINL(3)) & "'"
'            If Not rsprd.EOF Then ls.ListSubItems.Add , , Trim(rsprd(0))
'            rsact.Find "act_idy='" & Trim(RSINL(3)) & "'"
'            If Not rsact.EOF Then ls.ListSubItems.Add , , rsact(1)
'            ls.ListSubItems.Add , , RSINL(1)
'            RSINL.MoveNext
'        Wend
'        SSTab1.TabEnabled(3) = True
'        LVP.Visible = True
'        LVP.SetFocus
'        SSTab1.Tab = 3
'    Else
'        MsgBox "No Invoices Found", vbInformation + vbOKOnly, "DAS Version 4.0"
'        SSTab1.Tab = 6
'        ChrDate.SetFocus
'        SSTab1.TabEnabled(3) = False
'    End If
    If rsAvt.State = 1 Then rsAvt.Close
    Set rsAvt = Nothing
    rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' and avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') and avt_dat=ctod('" & ChrDate & "') order by mbl_idy", con, adOpenKeyset, adLockPessimistic
    If rsAvt.RecordCount > 0 Then
        MSR.Rows = 1
        MSR.Cols = 6
        MSR.TextMatrix(0, 0) = "Mobile ID"
        MSR.ColWidth(0) = 1300
        MSR.ColAlignment(0) = 1
        MSR.TextMatrix(0, 1) = "Invoice ID"
        MSR.ColWidth(1) = 1300
        MSR.TextMatrix(0, 2) = "Status"
        MSR.ColWidth(2) = 800
        MSR.TextMatrix(0, 3) = "SKU Name"
        MSR.ColWidth(3) = 2200
        MSR.TextMatrix(0, 4) = "Party"
        MSR.ColWidth(4) = 1740
        MSR.TextMatrix(0, 5) = "Invoice Date"
        MSR.ColWidth(5) = 1200
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
        MSR.Rows = rsAvt.RecordCount + 1
'                MSA.FixedRows = 1
        r = 1
        While Not rsAvt.EOF
            MSR.TextMatrix(r, 0) = Trim(rsAvt(0))
            MSR.TextMatrix(r, 1) = Trim(rsAvt(1))
            MSR.TextMatrix(r, 2) = "No" 'Trim(rsavt(6)) & ""
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsprd.EOF Then MSR.TextMatrix(r, 3) = Trim(rsprd(0))
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act where act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then MSR.TextMatrix(r, 4) = Trim(rsact(1))
            MSR.TextMatrix(r, 5) = Trim(rsAvt(2))
            r = r + 1
            rsAvt.MoveNext
        Wend
        SSTab1.Tab = 6
        LVP.Visible = False
        SSTab1.TabEnabled(3) = False
        MSR.Col = 2
        MSR.SetFocus
    Else
        MsgBox "No Invocies found in the selected Date!", vbInformation + vbOKOnly, "DAS Version 4.0"
        ChrDate.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub cmbType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 
    If cmbType.ListIndex = 0 Then
        'modified by parvathi
        Label6.Visible = True
        txtRet.Visible = True
        txtRetName.Visible = True
        txtRet.Text = ""
        txtRetName.Caption = ""
        'modified by parvathi
        LVF.Visible = True
        LVR.Visible = True
        Command5.Enabled = True
        Command6.Enabled = True
        MSPC.Visible = False
        txtRet.Enabled = True
        Label9.Visible = True
        txtRet.SetFocus
    Else
        'modified by parvathi
        Label6.Visible = False
        txtRet.Visible = False
        txtRetName.Visible = False
        'modified by parvathi
        MSPC.Visible = True
        LVF.Visible = False
        LVR.Visible = False
        Command5.Enabled = False
        Command6.Enabled = False
        dispPurInv
        Label9.Visible = False
        MSPC.Col = 3
        MSPC.SetFocus
    End If
End If
End Sub

Private Sub CMDAUPDATE_Click()
Dim RST_CK1 As New ADODB.Recordset
If MSA.Rows = 1 Then MSA.Rows = 2
If Len(Trim(MSA.TextMatrix(1, 1))) > 0 Then
If MsgBox("Update Transactions?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    clrall
    Exit Sub
End If
If rsAvt.State = 1 Then rsAvt.Close
Set rsAvt = Nothing
rsAvt.Open "select MBL_IDY,AVT_STS,AVT_DAT,AVT_TIM,SAL_IDY,MBL_ODY from avt where sal_typ='SAL' and (avt_sts is null or avt_sts='' or avt_sts<>'Yes')", con, adOpenKeyset, adLockPessimistic
'rsAvt.Open "SELECT MBL_IDY,AVT_STS,AVT_DAT,AVT_TIM,SAL_IDY FROM AVT", CON, adOpenKeyset, adlockpessimistic
For r = 1 To MSA.Rows - 1
    If Trim(MSA.TextMatrix(r, 3)) = "Yes" Or Trim(MSA.TextMatrix(r, 3)) = "yes" Then
        If rsAvt.State = 1 Then rsAvt.Close
        'rsAvt.Open "SELECT MBL_IDY,AVT_STS,AVT_DAT,AVT_TIM,SAL_IDY FROM AVT", CON, adOpenKeyset, adlockpessimistic
        rsAvt.Open "select MBL_IDY,AVT_STS,AVT_DAT,AVT_TIM,SAL_IDY,MBL_ODY from avt where sal_typ='SAL' and MBL_IDY='" & Trim(MSA.TextMatrix(r, 1)) & "' AND (avt_sts is NULL or avt_sts='' or avt_sts<>'Yes')", con, adOpenKeyset, adLockPessimistic
        If Not rsAvt.EOF Then
            rsAvt("AVT_STS") = Trim(MSA.TextMatrix(r, 3)) & ""
            rsAvt("AVT_DAT") = avtdate
            If MSA.Cols = 8 Then
            If Len(Trim(Val(MSA.TextMatrix(r, 7)))) = 10 Then
                If RST_CK1.State = 1 Then RST_CK1.Close
                RST_CK1.Open "select MBL_IDY,AVT_STS,AVT_DAT,AVT_TIM,SAL_IDY,MBL_ODY from avt where sal_typ='SAL' and MBL_IDY='" & Trim(MSA.TextMatrix(r, 7)) & "' AND (avt_sts is null or avt_sts='' or avt_sts<>'Yes')", con, adOpenKeyset, adLockPessimistic
                    If RST_CK1.EOF Then
                        rsAvt("MBL_IDY") = Trim(MSA.TextMatrix(r, 7))
                        rsAvt("MBL_ODY") = Trim(MSA.TextMatrix(r, 1))
                    Else
                        rsAvt.CancelUpdate
                        GoTo AA
                    End If
                End If
            End If
            rsAvt("AVT_TIM") = Time$
            rsAvt.Update
AA:
        End If
    End If
Next
clrall
Else
    MsgBox "Invoice not Selected!", vbInformation, "DAS Version 4.0"
End If
End Sub

Private Sub cmdCUpdate_Click()
Dim i  As Integer
If MSC.Rows = 1 Then MSC.Rows = 2
If Len(Trim(MSC.TextMatrix(1, 1))) > 0 Then
If MsgBox("Save CAF - Customer Agreement Forms Receipt Confirmation ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    CLRCLL
    Exit Sub
End If
If rsAvt.State = 1 Then rsAvt.Close
rsAvt.Open "SELECT MBL_IDY,caf_STS,AVT_DAT,AVT_TIM,SAL_IDY FROM AVT ", con, adOpenKeyset, adLockPessimistic
For r = 1 To MSC.Rows - 1
    If rsAvt.State = 1 Then rsAvt.Close
    rsAvt.Open "SELECT MBL_IDY,caf_STS,AVT_DAT,AVT_TIM,SAL_IDY FROM AVT where MBL_IDY='" & Trim(MSC.TextMatrix(r, 1)) & "' and avt_sts='Yes' AND (CHR_STS<>'Yes' or chr_sts='No' or chr_sts is null)", con, adOpenKeyset, adLockPessimistic
    i = rsAvt.RecordCount
    i = 0
    If Not rsAvt.EOF Then
        rsAvt("caf_STS") = Trim(MSC.TextMatrix(r, 3)) & ""
        rsAvt.Update
    End If
Next
CLRCLL
Else
    MsgBox "Invoice not Selected!", vbInformation, "DAS Version 4.0"
End If
End Sub

Private Sub cmdModUpdate_Click()
Dim fff As Boolean
Dim mbl_lst() As String
Dim rst_upd As New ADODB.Recordset
Dim rst_upd1 As New ADODB.Recordset
Dim pr As Integer
Dim SET_NUM As Integer

If SRT = "SAL" Then

fff = False
SET_NUM = 0
pr = 0

While pr <= Val(TT.Caption)
    fff = False
    If rsMod.State = 1 Then rsMod.Close
    
    rsMod.Open "select SUM(INV.PRD_QTY),INV.prd_idy FROM INV,PRD WHERE PRD.EXC_INF IN('y','Y') and INV_IDY='" & Trim(txtModInv) & "'AND INV.prd_idy='" & Trim(PRDID(pr).Caption) & "' AND INV.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenStatic
    If rst_upd.State = 1 Then rst_upd.Close
    rst_upd.Open "select count(prd_idy) from avt where sal_dat=ctod('" & modDate.Value & "') and sal_idy='" & Trim(txtModInv) & "' and prd_idy='" & Trim(PRDID(pr).Caption) & "'", con, adOpenKeyset, adLockPessimistic
                
    If rst_upd.EOF = False And rsMod.EOF = False Then
        
        If rsMod(0) <> rst_upd(0) Then
                      
           SET_NUM = 0

            For i = 1 To LVMOD.ListItems.count
            
                If rst_upd1.State = 1 Then rst_upd1.Close
                rst_upd1.Open "select MBL_idy from avt where sal_dat=ctod('" & modDate.Value & "') and sal_idy='" & Trim(txtModInv) & "' and prd_idy='" & Trim(PRDID(pr).Caption) & "' AND MBL_IDY='" & Trim(LVMOD.ListItems(i)) & "'", con, adOpenKeyset, adLockPessimistic
                If rst_upd1.EOF = False Then
                If LVMOD.ListItems(i).Checked = False Then
                SET_NUM = Val(SET_NUM) + 1
                End If
                End If
            
            Next
                
            If rst_upd(0) - Val(SET_NUM) = rsMod(0) Then
                fff = True
            Else
                MsgBox "Retail Returns DO NOT MATCH!", vbInformation, "DAS Version 4.0"
                LVMOD.SetFocus
                Exit Sub
            End If
                
                
        End If
    End If
 pr = Val(pr) + 1
    
Wend

SET_NUM = 0
pr = 0

While pr <= Val(TT.Caption)
    fff = False
    If rsMod.State = 1 Then rsMod.Close
    
    rsMod.Open "select SUM(INV.PRD_QTY),INV.prd_idy FROM INV,PRD WHERE PRD.EXC_INF IN('y','Y') and INV_IDY='" & Trim(txtModInv) & "'AND INV.prd_idy='" & Trim(PRDID(pr).Caption) & "' AND INV.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenStatic
    If rst_upd.State = 1 Then rst_upd.Close
    rst_upd.Open "select count(prd_idy) from avt where sal_dat=ctod('" & modDate.Value & "') and sal_idy='" & Trim(txtModInv) & "' and prd_idy='" & Trim(PRDID(pr).Caption) & "'", con, adOpenKeyset, adLockPessimistic
                
    If rst_upd.EOF = False And rsMod.EOF = False Then
        
        If rsMod(0) <> rst_upd(0) Then
                      
           SET_NUM = 0

            For i = 1 To LVMOD.ListItems.count
            
                If rst_upd1.State = 1 Then rst_upd1.Close
                rst_upd1.Open "select MBL_idy from avt where sal_dat=ctod('" & modDate.Value & "') and sal_idy='" & Trim(txtModInv) & "' and prd_idy='" & Trim(PRDID(pr).Caption) & "' AND MBL_IDY='" & Trim(LVMOD.ListItems(i)) & "'", con, adOpenKeyset, adLockPessimistic
                If rst_upd1.EOF = False Then
                If LVMOD.ListItems(i).Checked = False Then
                SET_NUM = Val(SET_NUM) + 1
                End If
                End If
            
            Next
                
            If rst_upd(0) - Val(SET_NUM) = rsMod(0) Then
                
                For i = 1 To LVMOD.ListItems.count
            
                    If rst_upd1.State = 1 Then rst_upd1.Close
                    rst_upd1.Open "select MBL_idy from avt where sal_dat=ctod('" & modDate.Value & "') and sal_idy='" & Trim(txtModInv) & "' and prd_idy='" & Trim(PRDID(pr).Caption) & "' AND MBL_IDY='" & Trim(LVMOD.ListItems(i)) & "'", con, adOpenKeyset, adLockPessimistic
                    If rst_upd1.EOF = False Then
                    If LVMOD.ListItems(i).Checked = False Then
                        If rsaMod.State = 1 Then rsaMod.Close
                        Set rsaMod = Nothing
                        rsaMod.Open "select * from avt where mbl_idy='" & Trim(LVMOD.ListItems(i).Text) & "'", con, adOpenKeyset, adLockPessimistic
                        If Not rsaMod.EOF Then
                            rsaMod("smn_idy") = ""
                            rsaMod("sal_idy") = ""
                            rsaMod("sal_dat") = Null
                            rsaMod("cst_idy") = ""
                            rsaMod("sal_typ") = ""
                            rsaMod.Update
                        End If
                    End If
                    End If
                
                Next
                fff = True
            Else
                MsgBox "Retail Returns DO NOT MATCH!", vbInformation, "DAS Version 4.0"
                LVMOD.SetFocus
                Exit Sub
            End If
                
                
        End If
    End If
 pr = Val(pr) + 1
    
Wend

If fff = True Then
    lblno(0).Visible = False
    lblno(3).Visible = False
    lblno(2).Visible = False
    lblno(1).Visible = False
    MsgBox "Retail Returns Save Confirmed!", vbInformation + vbOKOnly, "DAS Version 4.0"
    LVMOD.ListItems.CLEAR
    txtModInv = ""
    txtModRet = ""
    txtModRetName = ""
    modDate = Date$
    modDate.SetFocus
End If
Else
'-----------------------------------------------------------------------
'Dim fff As Boolean
'Dim mbl_lst() As String
'Dim rst_upd As New ADODB.Recordset
'Dim rst_upd1 As New ADODB.Recordset
'Dim pr As Integer
'Dim SET_NUM As Integer

fff = False
SET_NUM = 0
pr = 0

While pr <= Val(TT.Caption)
    fff = False
    If rsMod.State = 1 Then rsMod.Close
    
    rsMod.Open "select SUM(INV.PRD_QTY),INV.prd_idy FROM INV,PRD WHERE PRD.EXC_INF IN('y','Y') and INV_IDY='" & Trim(txtModInv) & "'AND INV.prd_idy='" & Trim(PRDID(pr).Caption) & "' AND INV.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenStatic
    If rst_upd.State = 1 Then rst_upd.Close
    
    rst_upd.Open "select count(prd_idy) from avt where INV_dat=ctod('" & modDate.Value & "') and INV_idy='" & Trim(txtModInv) & "' and prd_idy='" & Trim(PRDID(pr).Caption) & "'", con, adOpenKeyset, adLockPessimistic
                
    If rst_upd.EOF = False And rsMod.EOF = False Then
        
        If rsMod(0) <> rst_upd(0) Then
                      
           SET_NUM = 0

            For i = 1 To LVMOD.ListItems.count
            
                If rst_upd1.State = 1 Then rst_upd1.Close
                rst_upd1.Open "select MBL_idy from avt where INV_dat=ctod('" & modDate.Value & "') and INV_idy='" & Trim(txtModInv) & "' and prd_idy='" & Trim(PRDID(pr).Caption) & "' AND MBL_IDY='" & Trim(LVMOD.ListItems(i)) & "'", con, adOpenKeyset, adLockPessimistic
                If rst_upd1.EOF = False Then
                If LVMOD.ListItems(i).Checked = False Then
                SET_NUM = Val(SET_NUM) + 1
                End If
                End If
            
            Next
                
            If rst_upd(0) - Val(SET_NUM) = rsMod(0) Then
                fff = True
            Else
                MsgBox "Retail Returns DO NOT MATCH!", vbInformation, "DAS Version 4.0"
                LVMOD.SetFocus
                Exit Sub
            End If
                
                
        End If
    End If
 pr = Val(pr) + 1
    
Wend

SET_NUM = 0
pr = 0

While pr <= Val(TT.Caption)
    fff = False
    If rsMod.State = 1 Then rsMod.Close
    
    rsMod.Open "select SUM(INV.PRD_QTY),INV.prd_idy FROM INV,PRD WHERE PRD.EXC_INF IN('y','Y') and INV_IDY='" & Trim(txtModInv) & "'AND INV.prd_idy='" & Trim(PRDID(pr).Caption) & "' AND INV.PRD_IDY LIKE PRD.PRD_IDY ", con, adOpenStatic
    If rst_upd.State = 1 Then rst_upd.Close
    rst_upd.Open "select count(prd_idy) from avt where INV_dat=ctod('" & modDate.Value & "') and INV_idy='" & Trim(txtModInv) & "' and prd_idy='" & Trim(PRDID(pr).Caption) & "'", con, adOpenKeyset, adLockPessimistic
                
    If rst_upd.EOF = False And rsMod.EOF = False Then
        
        If rsMod(0) <> rst_upd(0) Then
                      
           SET_NUM = 0

            For i = 1 To LVMOD.ListItems.count
            
                If rst_upd1.State = 1 Then rst_upd1.Close
                rst_upd1.Open "select MBL_idy from avt where INV_dat=ctod('" & modDate.Value & "') and INV_idy='" & Trim(txtModInv) & "' and prd_idy='" & Trim(PRDID(pr).Caption) & "' AND MBL_IDY='" & Trim(LVMOD.ListItems(i)) & "'", con, adOpenKeyset, adLockPessimistic
                If rst_upd1.EOF = False Then
                If LVMOD.ListItems(i).Checked = False Then
                SET_NUM = Val(SET_NUM) + 1
                End If
                End If
            
            Next
                
            If rst_upd(0) - Val(SET_NUM) = rsMod(0) Then
                
                For i = 1 To LVMOD.ListItems.count
            
                    If rst_upd1.State = 1 Then rst_upd1.Close
                    rst_upd1.Open "select MBL_idy from avt where INV_dat=ctod('" & modDate.Value & "') and INV_idy='" & Trim(txtModInv) & "' and prd_idy='" & Trim(PRDID(pr).Caption) & "' AND MBL_IDY='" & Trim(LVMOD.ListItems(i)) & "'", con, adOpenKeyset, adLockPessimistic
                    If rst_upd1.EOF = False Then
                    If LVMOD.ListItems(i).Checked = False Then
                        If rsaMod.State = 1 Then rsaMod.Close
                        Set rsaMod = Nothing
                        rsaMod.Open "select * from avt where mbl_idy='" & Trim(LVMOD.ListItems(i).Text) & "'", con, adOpenKeyset, adLockPessimistic
                        If Not rsaMod.EOF Then
                            con.Execute "DELETE FROM AVT WHERE mbl_idy='" & Trim(LVMOD.ListItems(i).Text) & "'"
                        End If
                    End If
                    End If
                
                Next
                fff = True
            Else
                MsgBox "Retail Returns DO NOT MATCH!", vbInformation, "DAS Version 4.0"
                LVMOD.SetFocus
                Exit Sub
            End If
                
                
        End If
    End If
 pr = Val(pr) + 1
    
Wend

If fff = True Then
    lblno(0).Visible = False
    lblno(3).Visible = False
    lblno(2).Visible = False
    lblno(1).Visible = False
    MsgBox "Purchase Returns Save Confirmed!", vbInformation + vbOKOnly, "DAS Version 4.0"
    LVMOD.ListItems.CLEAR
    txtModInv = ""
    txtModRet = ""
    txtModRetName = ""
    modDate = Date$
    modDate.SetFocus
End If
'-----------------------------------------------------------------------
End If
SRT = ""
End Sub

Private Sub cmdRUpdate_Click()
If MSR.Rows = 1 Then MSR.Rows = 2
If Len(Trim(MSR.TextMatrix(1, 1))) > 0 Then
If MsgBox("Update Transactions?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    CLRRLL
    Exit Sub
End If
If rsAvt.State = 1 Then rsAvt.Close
rsAvt.Open "SELECT MBL_IDY,chr_STS,chr_dat,AVT_DAT,AVT_TIM,SAL_IDY FROM AVT ", con, adOpenKeyset, adLockPessimistic
For r = 1 To MSR.Rows - 1
    If Trim(MSR.TextMatrix(r, 3)) = "Yes" Or Trim(MSR.TextMatrix(r, 3)) = "yes" Then
    If rsAvt.State = 1 Then rsAvt.Close
    rsAvt.Open "SELECT MBL_IDY,chr_STS,chr_dat,AVT_DAT,AVT_TIM,SAL_IDY FROM AVT WHERE MBL_IDY='" & Trim(MSR.TextMatrix(r, 1)) & "' AND AVT_STS='Yes'AND (CHR_STS<>'Yes' or chr_sts='No' or chr_sts is null)", con, adOpenKeyset, adLockPessimistic
    If Not rsAvt.EOF Then
        rsAvt("chr_STS") = Trim(MSR.TextMatrix(r, 3)) & ""
        rsAvt("chr_dat") = avtcdate
        rsAvt.Update
    End If
    End If
Next
CLRRLL
Else
    MsgBox "Invoice not Selected!", vbInformation, "DAS Version 4.0"
End If
End Sub

Private Sub cmdsUpdate_Click()

Dim sav_rst As New ADODB.Recordset

If Len(Trim(txtSalInv)) > 0 Then
If MsgBox("Update Transactions?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    lv.ListItems.CLEAR
    sinvdat = Date$
    txtSalInv = ""
    txtPartyname = ""
    txtSmn = ""
    txtMarket = ""
    txtSMbl = ""
    txtSalInv.SetFocus
    Exit Sub
End If

If lv.ListItems.count = 0 Then
    MsgBox "No Invoice Found", vbInformation + vbOKOnly, "DAS Version 4.0"
    txtSalInv.SetFocus
    Exit Sub
End If
Dim cn As Integer
If RSINV.State = 1 Then RSINV.Close
Set RSINV = Nothing
RSINV.Open "SELECT sum(iNV.PRD_QTY),inv.prd_idy,prd.prd_nme,exc_inf FROM INV,prd WHERE INV_IDY='" & Trim(txtSalInv) & "' and inv.prd_idy=prd.prd_idy and exc_inf='Y' and inv.stk_typ in('Salable','Free') GROUP by inv.prd_idy", con, adOpenStatic
If rschk.State = 1 Then rschk.Close
Set rschk = Nothing
rschk.Open "SELECT PRD_IDY FROM INV WHERE INV_IDY='" & Trim(txtSalInv) & "' GROUP BY PRD_IDY", con, adOpenKeyset, adLockPessimistic
If rsAvt.State = 1 Then rsAvt.Close
Set rsAvt = Nothing
rsAvt.Open "SELECT MBL_IDY,SAL_IDY,SAL_DAT,SAL_TYP,CST_IDY,SMN_IDY FROM AVT WHERE (CHR_STS<>'Yes' OR chr_sts='No' or chr_sts is null)", con, adOpenKeyset, adLockPessimistic
Dim mblc As Integer
mblc = 0
While Not RSINV.EOF
    cn = 0
    For r = 1 To lv.ListItems.count
    'MsgBox RSINV(2)
        If lv.ListItems(r).Checked = True And Trim(lv.ListItems(r).ListSubItems(5).Text) = Trim(RSINV(1)) And Trim(lv.ListItems(r).ListSubItems(1).Text) = Trim(RSINV(2)) Then
            cn = cn + 1
        End If
    Next
    If sav_rst.State = 1 Then sav_rst.Close
    sav_rst.Open "select prd_idy from avt where sal_dat=ctod('" & sinvdat.Value & "') and sal_idy='" & Trim(txtSalInv) & "' and prd_idy='" & Trim(RSINV(1)) & "'", con, adOpenKeyset, adLockPessimistic
    
         If sav_rst.EOF = False Then
         mblc = Val(sav_rst.RecordCount)
         End If
         cn = Val(cn) + Val(mblc)
        If cn <= RSINV(0) Then
        For r = 1 To lv.ListItems.count
            rsAvt.MoveFirst
            If lv.ListItems(r).Checked = True And Trim(lv.ListItems(r).ListSubItems(5).Text) = Trim(RSINV(1)) And Trim(lv.ListItems(r).ListSubItems(1).Text) = Trim(RSINV(2)) Then
                rsAvt.Find "mbl_idy='" & Trim(lv.ListItems(r).Text) & "'" 'inv_dat=ctod('" & sinvdat & "')"
                If Not rsAvt.EOF Then
                rsAvt("sal_idy") = Trim(txtSalInv)
                rsAvt("sal_dat") = sinvdat.Value
                rsAvt("sal_typ") = TTYP
                rsAvt("cst_idy") = Trim(lv.ListItems(r).ListSubItems(4).Text) & ""
                rsAvt("smn_idy") = Trim(lv.ListItems(r).ListSubItems(6).Text) & ""
                rsAvt.Update
            End If
            End If
            Next
        Else
            For r = 1 To lv.ListItems.count
            rsAvt.MoveFirst
            If lv.ListItems(r).Checked = True And Trim(lv.ListItems(r).ListSubItems(5).Text) = Trim(RSINV(1)) And Trim(lv.ListItems(r).ListSubItems(1).Text) = Trim(RSINV(2)) Then
                rsAvt.Find "mbl_idy='" & Trim(lv.ListItems(r).Text) & "'" 'inv_dat=ctod('" & sinvdat & "')"
                If Not rsAvt.EOF Then
                rsAvt("sal_idy") = Trim(txtSalInv)
                rsAvt("sal_dat") = sinvdat.Value
                rsAvt("sal_typ") = TTYP
                rsAvt("cst_idy") = Trim(lv.ListItems(r).ListSubItems(4).Text) & ""
                rsAvt("smn_idy") = Trim(lv.ListItems(r).ListSubItems(6).Text) & ""
                rsAvt.Update
            End If
            End If
            Next
            For s = 0 To 5
                If Ptno(s).Visible = True Then Ptno(s).Visible = False
            Next
        End If

    RSINV.MoveNext
Wend
    For i = 0 To Val(pn.Caption)
    Ptno(i).Visible = False
    Next
    lv.ListItems.CLEAR
    txtSalInv = ""
    txtPartyname = ""
    sinvdat = Date$
    txtSmn = ""
    txtMarket = ""
    txtSMbl = ""
    txtSalInv.SetFocus
Else
     MsgBox "Invoice not Selected!", vbInformation, "DAS Version 4.0"
End If
End Sub
Private Sub Command1_Click()
Dim RSINV1 As New ADODB.Recordset

Dim sup_chk As Boolean
sup_chk = False
On Error GoTo WS
If rschk.State = 1 Then rschk.Close
If RSINV.State = 1 Then RSINV.Close
RSINV.Open "select inv.inv_idy,inv.inv_dat,inv.prd_idy,inv.prd_qty,cst_idy,inv.txn_typ,exc_inf from inv,inl,prd where inv.inv_dat=inl.inv_dat and inv.inv_idy like inl.inv_idy and inv.txn_typ in('PUR','OPS','STI','SRN') and inv.prd_idy like prd.prd_idy and prd.exc_inf='Y' order by inv.inv_idy", con, adOpenKeyset, adLockPessimistic
'If RSINV.RecordCount = 0 Then
'MsgBox "Invoice not Found!", vbInformation, "DAS Version 4.0"
'Exit Sub
'End If
'RSINV.Open "select inv.inv_idy,inv.inv_dat,inv.prd_idy,inv.prd_qty,cst_idy,inv.txn_typ,exc_inf from inv,inl,prd where inv.inv_dat=inl.inv_dat and inv.inv_idy like inl.inv_idy and inv.txn_typ in('PUR','OPS','STI','SRN') and inv.prd_idy like prd.prd_idy and prd.exc_inf='Y' order by inv.inv_idy", CON, adOpenKeyset, adlockpessimistic
While Not RSINV.EOF
    If rschk.State = 1 Then rschk.Close
    rschk.Open "Select alt_idy from prd where prd_idy ='" & RSINV(2) & "'", con, adOpenKeyset, adLockPessimistic
    If rschk.EOF = False Then
        If IsNull(rschk(0)) = True Or Trim(rschk(0)) = "" Then
            sup_chk = True
        End If
    End If
RSINV.MoveNext
Wend
If sup_chk = False Then
RSINV.MoveFirst
If rsact.State = 1 Then rsact.Close
While Not RSINV.EOF
    If rschk.State = 1 Then rschk.Close
    rschk.Open "select AVT.inv_idy, AVT.inv_dat, AVT.prd_idy from avt order by inv_idy", con, adOpenKeyset, adLockPessimistic
    If rschk.RecordCount > 0 Then rschk.MoveFirst
    rschk.filter = "inv_idy='" & Trim(RSINV(0)) & "' and inv_dat=#" & RSINV(1) & "# and prd_idy='" & Trim(RSINV(2)) & "'"
    If rschk.EOF Then   ' checking whether invoice exists or not
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select grp_idy from prd where prd_idy like '" & Trim(RSINV(2)) & "'", con, adOpenKeyset, adLockPessimistic
        If RSINV(3) > 0 Then    ' If Trans Qty is greater than 0
            pqty = RSINV(3)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select * from avt", con, adOpenKeyset, adLockPessimistic
            For r = 1 To pqty
                rsAvt.AddNew
                rsAvt("inv_idy") = Trim(RSINV(0)) & ""
                rsAvt("inv_dat") = RSINV(1)
                rsAvt("prd_idy") = Trim(RSINV(2)) & ""
                If Not rsprd.EOF Then rsAvt("cat_idy") = rsprd(0)
                If Trim(RSINV(5)) = "PUR" Or Trim(RSINV(5)) = "OPS" Or Trim(RSINV(5)) = "STI" Or Trim(RSINV(5)) = "SRN" Then
                    rsAvt("SPL_IDY") = Trim(RSINV(4)) & ""
                Else
                    rsAvt("CST_IDY") = Trim(RSINV(4)) & ""
                End If
                rsAvt("txn_typ") = RSINV(5)
                rsAvt.Update
            Next
        End If
    End If
    RSINV.MoveNext
Wend
MsgBox "Mobile Number Records Created Successfully!", vbInformation + vbOKOnly, "DAS Version 4.0"
txtVch.Enabled = True
txtVch.SetFocus
Else
    MsgBox "Supplier's Product Code not Specified to the Product!", vbCritical, "DAS Version 4.0"
End If
Exit Sub
WS:
txtVch.Enabled = True
txtVch.SetFocus
'MsgBox err.Description, vbInformation + vbOKOnly, "DAS Version 4.0"
End Sub

Private Sub CmdGEN_Click()
Dim gen_rst As New ADODB.Recordset
Dim ppp As Integer
Dim gen_num As Integer
gen_num = 0
ppp = 1
If Len(Trim(Val(txtStart))) = 10 And txtTemp.Visible = False Then
   
    While ppp <= MS.Rows - 1
        gen_rst.Open "select mbl_idy from avt where mbl_idy='" & Trim(Val(txtStart) + gen_num) & "' AND (chr_sts is null or chr_sts='' or chr_sts<>'Yes')", con, adOpenStatic
        If gen_rst.EOF Then
        MS.TextMatrix(ppp, 3) = Val(txtStart) + gen_num
        ppp = ppp + 1
        End If
        gen_num = gen_num + 1
        gen_rst.Close
    Wend
    MS.Row = 1
    MS.Col = 3
    MS.SetFocus
ElseIf Len(Trim(Val(txtTemp))) = 10 Then  ' ADDED BY PKB FOR BOTH AUTO AND WITHOUT AUTO GENERATION
   txtStart = ""
   ppp = MS.Row
    While ppp <= MS.Rows - 1
    If gen_rst.State = 1 Then gen_rst.Close
        gen_rst.Open "select mbl_idy from avt where mbl_idy='" & Trim(Val(txtTemp) + gen_num) & "' AND (chr_sts is null or chr_sts='' or chr_sts<>'Yes')", con, adOpenStatic
        If gen_rst.EOF Then
        MS.TextMatrix(ppp, 3) = Val(txtTemp) + gen_num
        ppp = ppp + 1
        End If
        gen_num = gen_num + 1
        gen_rst.Close
    Wend
    MS.Row = 1
    MS.Col = 3
    MS.SetFocus
    txtTemp.Visible = False

Else
    MsgBox "10 Digit Mobile Number Required!", vbOKOnly, "DAS Version 4.0"
    txtStart.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub Command10_Click()

End Sub


Private Sub Command3_Click()
updateavt
inward.Visible = False
End Sub

Private Sub Command4_Click()

If MSA.Rows = 1 Then MSA.Rows = 2

Dim RSTEMP As New ADODB.Recordset
Dim RSIMP1 As New ADODB.Recordset
Dim appXL As Object
Dim appX As Object
Dim TYP As String
Dim adate As Date

Set appXL = CreateObject("Excel.Application")

CommonDialog1.filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
CommonDialog1.DialogTitle = "Select Mobile Numbers File"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen

CommonDialog1.CancelError = False
If Not Trim(CommonDialog1.FileName) = "" Then
    If LCase(Right(Trim(CommonDialog1.FileName), 3)) = "xls" Then ' Edited
         appXL.Workbooks.Open (CommonDialog1.FileName)  'GetObject(Trim(CommonDialog1.FileName))
    Else
        MsgBox "Selected File is not a Valid SpreadSheet File"
        Exit Sub
    End If
    
    TYP = appXL.Worksheets(1).Cells(1, 1)
    adate = appXL.Worksheets(1).Cells(2, 1)
    
    If TYP = "Churn" Then
    
        If Val(Mid(adate, 1, 2)) <= 12 And Trim(Mid(adate, 3, 1)) = "/" And Trim(Mid(adate, 6, 1)) = "/" Then
            
            avtcdate.Value = Trim(adate)
        
            Dim rstemp1 As New ADODB.Recordset
            Dim rstemp2 As New ADODB.Recordset
            Dim pno1 As Integer
        
            On Error GoTo errh
            Dim mn() As String
            Dim mn1() As String
        
            Dim chk_flg As Boolean
            chk_flg = False
        
            Dim pno As Integer
            Dim row_num As Integer
            
            con.Execute "delete from rp2", qc
            
            If RSTEMP.State = 1 Then RSTEMP.Close
            RSTEMP.Open "select chr_one,chr_two from rp2", con, adOpenKeyset, adLockPessimistic
            
            row_num = 4
            tot_num = 0
            
            pno = 0
            While Len(Trim(appXL.Worksheets(1).Cells(row_num, 1))) > 0
                If Len(Trim(Val(appXL.Worksheets(1).Cells(row_num, 1)))) = 10 Then
                    If RSTEMP.State = 1 Then RSTEMP.Close
                    RSTEMP.Open "select chr_one,chr_two from rp2 WHERE CHR_ONE='" & Trim(appXL.Worksheets(1).Cells(row_num, 1)) & "'", con, adOpenKeyset, adLockPessimistic
                    If RSTEMP.EOF Then
                       RSTEMP.AddNew
                       RSTEMP(0) = appXL.Worksheets(1).Cells(row_num, 1)
                       RSTEMP.Update
                       tot_num = Val(tot_num) + 1
                     End If
                 End If
                 row_num = Val(row_num) + 1
            Wend
            
            For i = 1 To appXL.Workbooks.count
            appXL.Workbooks(i).Close
            Next
            Set appXL = Nothing
        
        '        Label1(8).Caption = Label1(8).Caption & ":"
            If RSINL.State = 1 Then RSINL.Close
            MSR.Rows = 1
            MSR.Cols = 7
            RSINL.Open "select mbl_idy,sal_idy,inv_dat,prd_idy,cst_idy,txn_typ,avt_sts,caf_sts from avt where MBL_IDY IN (SELECT CHR_ONE FROM RP2) AND sal_typ='SAL' AND avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='')", con, adOpenKeyset, adLockPessimistic
            If RSINL.RecordCount > 0 Then
                MSR.Rows = 1
                MSR.Cols = 7
                MSR.TextMatrix(0, 0) = "SNO"
                MSR.ColWidth(0) = 700
                MSR.TextMatrix(0, 1) = "Mobile ID"
                MSR.ColWidth(1) = 1500
                MSR.ColAlignment(1) = 1
                MSR.TextMatrix(0, 2) = "Invoice ID"
                MSR.ColWidth(2) = 1600
                MSR.TextMatrix(0, 3) = "Status"
                MSR.ColWidth(3) = 800
                MSR.TextMatrix(0, 4) = "SKU Name"
                MSR.ColWidth(4) = 2000
                MSR.TextMatrix(0, 5) = "Party"
                MSR.ColWidth(5) = 2000
                MSR.TextMatrix(0, 6) = "Invoice Date"
                MSR.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSR.Rows = RSINL.RecordCount + 1
                r = 1
                While Not RSINL.EOF
                    MSR.TextMatrix(r, 0) = r
                    MSR.TextMatrix(r, 1) = Trim(RSINL(0))
                    If IsNull(RSINL(1)) = True Then
                    MSR.TextMatrix(r, 2) = " "
                    Else
                    MSR.TextMatrix(r, 2) = Trim(RSINL(1))
                    End If
                    MSR.TextMatrix(r, 3) = "Yes" 'Trim(RSINL(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSR.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act where act_idy='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSR.TextMatrix(r, 5) = Trim(rsact(1))
                    MSR.TextMatrix(r, 6) = Trim(RSINL(2))
                    r = r + 1
                    RSINL.MoveNext
                Wend
                cmdRUpdate.SetFocus
                con.Execute "delete from rp2", qc
                Exit Sub
        Else
            MsgBox "Mobile Numbers are not availble in the file!", vbInformation, "DAS Version 4.0"
            appXL.WORKBOOKWS.Close
            Set appXL = Nothing
            Exit Sub
        End If
            MsgBox "Valid Date Specified in the file!", vbInformation, "DAS Version 4.0"
            appXL.WORKBOOKWS.Close
            Set appXL = Nothing
            Exit Sub
        End If
        
    Else
        MsgBox "Select a Valid File!", vbInformation, "DAS Version 4.0"
        appXL.WORKBOOKWS.Close
        Set appXL = Nothing
        Exit Sub
    End If
    Exit Sub
End If
errh:
MsgBox "Select a Valid File!", vbInformation, "DAS Version 4.0"

End Sub

Private Sub Command5_Click()
For r = 1 To LVF.ListItems.count
    If LVF.ListItems(r).Checked = True Then
        LVR.SetFocus
        Exit Sub
    End If
Next
MsgBox "Select Numbers to be Replaced!", vbInformation + vbOKOnly, "DAS Version 4.0"
LVF.SetFocus
End Sub

Private Sub Command6_Click()
If MsgBox("Update Changes?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
    List6.CLEAR
    List7.CLEAR
    LIST8.CLEAR
    For r = 1 To LVF.ListItems.count
        If LVF.ListItems(r).Checked = True Then
            List6.AddItem Trim(LVF.ListItems(r).Text)
        End If
    Next
    For r = 1 To LVR.ListItems.count
        If LVR.ListItems(r).Checked = True Then
            List7.AddItem Trim(LVR.ListItems(r).Text)
            LIST8.AddItem Trim(LVR.ListItems(r).Text)
        End If
    Next
    If List6.ListCount > 0 And List7.ListCount > 0 Then
        If List6.ListCount = List7.ListCount Then
            
                For r = 0 To List6.ListCount - 1
                    If rsEdt.State = 1 Then rsEdt.Close
                    Set rsEdt = Nothing
                    rsEdt.Open "select * from avt where mbl_idy='" & Trim(List6.List(r)) & "' and sal_typ='SAL'", con, adOpenKeyset, adLockPessimistic
                    If Not rsEdt.EOF Then
                        rsEdt("mbl_idy") = Trim(List7.List(r))
                        rsEdt.Update
                    End If
                Next
                For r = 0 To List7.ListCount - 1
                    If rsEdt.State = 1 Then rsEdt.Close
                    Set rsEdt = Nothing
                    rsEdt.Open "select * from avt where mbl_idy='" & Trim(List7.List(r)) & "' and (sal_typ is null or sal_typ<>'SAL')", con, adOpenKeyset, adLockPessimistic
                    rsEdt("mbl_idy") = Trim(List6.List(r))
                    rsEdt.Update
                Next
                LVR.ListItems.CLEAR
                LVF.ListItems.CLEAR
                txtRet = ""
                txtRetName = ""
                cmbType.SetFocus
                txtRet.Enabled = True
            
        Else
            MsgBox "Quanity does not Match!", vbInformation + vbOKOnly, "DAS Version 4.0"
            LVR.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Select Mobile Numbers to be replaced!", vbInformation + vbOKOnly, "DAS Version 4.0"
        LVF.SetFocus
        Exit Sub
    End If
Else
    txtRet = ""
    txtRetName = ""
    txtRet.Enabled = True
    LVF.ListItems.CLEAR
    LVR.ListItems.CLEAR
    cmbType.SetFocus
    Exit Sub
End If
End Sub

Private Sub Command7_Click()
'If txtVch.Text <> "" Then
'Dim appXL As Object
'Dim appX As Object
'Dim tot_num As Integer
'
'Set appXL = CreateObject("Excel.Application")
'
'CommonDialog1.filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
'CommonDialog1.DialogTitle = "Select Mobile Numbers File"
'CommonDialog1.FilterIndex = 2
'CommonDialog1.ShowOpen
'
'CommonDialog1.CancelError = False
'If Not Trim(CommonDialog1.FileName) = "" Then
'    If LCase(Right(Trim(CommonDialog1.FileName), 3)) = "xls" Then ' Edited
'         appXL.WORKBOOKS.Open (CommonDialog1.FileName)  'GetObject(Trim(CommonDialog1.FileName))
'    Else
'        MsgBox "Selected File is not a Valid SpreadSheet File"
'        Exit Sub
'    End If
'tot_num = appXL.Worksheets(1).Cells(1, 1)
'Dim row_num As Integer
'row_num = 3
'tot_num = 0
'While Len(Trim(appXL.Worksheets(1).Cells(row_num, 1))) > 0
'If Len(Trim(appXL.Worksheets(1).Cells(row_num, 1))) = 10 Then
'tot_num = Val(tot_num) + 1
'End If
'row_num = Val(row_num) + 1
'Wend
'
'If tot_num = MS.Rows - 1 Then
'      For r = 1 To MS.Rows - 1
'         MS.TextMatrix(r, 3) = Trim(appXL.Worksheets(1).Cells(r + 2, 1))
'      Next
'      updateavt
'      For r = 1 To MS.Rows - 1
'         MS.TextMatrix(r, 3) = ""
'      Next
'Else
'    MsgBox "Invalid Mobile Numbers SpreadSheet File Selected!", vbCritical + vbOKOnly, "DAS Version 4.0"
'End If
'Else
'    MsgBox "File Selection Cancelled!", vbInformation + vbOKOnly, "DAS Version 4.0"
'    Command7.SetFocus
'End If
'appXL.WORKBOOKS.Close
'appXL.Quit
'Set appXL = Nothing
'Else
'    MsgBox "Invoice not Selected!", vbInformation, "DAS Version 4.0"
'End If
End Sub

Private Sub Command8_Click()
If Len(Trim(txtSalInv)) > 0 Then
Dim appXL As Object
Dim appX As Object
Dim tot_num As Integer

Set appXL = CreateObject("Excel.Application")

CommonDialog1.filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
CommonDialog1.DialogTitle = "Select Mobile Numbers File"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen

CommonDialog1.CancelError = False
If Not Trim(CommonDialog1.FileName) = "" Then
    If LCase(Right(Trim(CommonDialog1.FileName), 3)) = "xls" Then ' Edited
         appXL.Workbooks.Open (CommonDialog1.FileName)
    Else
        MsgBox "Selected File is not a Valid SpreadSheet File"
        Exit Sub
    End If
    
tot_num = appXL.Worksheets(1).Cells(1, 1)
Dim row_num As Integer
row_num = 3
tot_num = 0

While Len(Trim(appXL.Worksheets(1).Cells(row_num, 1))) > 0
    If Len(Trim(appXL.Worksheets(1).Cells(row_num, 1))) = 10 Then
        txtSMbl.Text = Trim(appXL.Worksheets(1).Cells(row_num, 1))
        txtSMbl_KeyDown 13, 0
    End If
    row_num = Val(row_num) + 1
Wend
      '-----------------------------------------------------------------
Else
    MsgBox "File Selection Cancelled!", vbInformation + vbOKOnly, "DAS Version 4.0"
    Command7.SetFocus
End If
appXL.Workbooks.Close
appXL.Quit
Set appXL = Nothing
Else
   MsgBox "Invoice not Selected!", vbInformation, "DAS Version 4.0"
End If
txtSMbl.Text = ""
End Sub

Private Sub Command9_Click()
'using excel
If MSA.Rows = 1 Then MSA.Rows = 2
Dim appXL As Object
Dim appX As Object
Dim tot_num As Integer

If Len(Trim(MSA.TextMatrix(1, 1))) > 0 Then
On Error GoTo errh
If rsimp.State = 1 Then rsimp.Close
Set rsimp = Nothing
rsimp.Open "select * from avt where sal_typ='SAL' and (avt_sts is null or avt_sts='' or avt_sts<>'Yes')", con, adOpenKeyset, adLockPessimistic

Set appXL = CreateObject("Excel.Application")

CommonDialog1.filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
CommonDialog1.DialogTitle = "Select Mobile Numbers File"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen

CommonDialog1.CancelError = False
If Not Trim(CommonDialog1.FileName) = "" Then
    If LCase(Right(Trim(CommonDialog1.FileName), 3)) = "xls" Then ' Edited
         appXL.Workbooks.Open (CommonDialog1.FileName)
    Else
        MsgBox "Selected File is not a Valid SpreadSheet File"
        Exit Sub
    End If
    
tot_num = appXL.Worksheets(1).Cells(1, 1)
Dim row_num As Integer
row_num = 3
tot_num = 0
If MSA.Rows = 1 Then MSA.Rows = 2
If Len(Trim(MSA.TextMatrix(1, 1))) > 0 Then
On Error GoTo errh
If rsimp.State = 1 Then rsimp.Close
Set rsimp = Nothing
rsimp.Open "select * from avt where sal_typ='SAL' and (avt_sts is null or avt_sts='' or avt_sts<>'Yes')", con, adOpenKeyset, adLockPessimistic
adate = Trim(appXL.Worksheets(1).Cells(1, 1))

While Len(Trim(appXL.Worksheets(1).Cells(row_num, 1))) > 0
    '-----------------------------------------------------
    
    SMBL = Trim(appXL.Worksheets(1).Cells(row_num, 1))
    
    If rsimp.RecordCount > 0 Then rsimp.MoveFirst
    rsimp.Find "mbl_idy='" & Trim(SMBL) & "'"
    If Not rsimp.EOF Then
        
        rsimp("avt_sts") = "Yes"
        rsimp("avt_dat") = CDate(adate)
        If Len(Trim(appXL.Worksheets(1).Cells(row_num, 2))) > 0 Then
            rsimp("MBL_ODY") = Trim(appXL.Worksheets(1).Cells(row_num, 2))
            rsimp("MBL_IDY") = Trim(appXL.Worksheets(1).Cells(row_num, 1))
        End If
        rsimp("avt_tim") = Time$
        rsimp.Update
    End If
    row_num = Val(row_num) + 1
Wend

appXL.Workbooks.Close
appXL.Quit
Set appXL = Nothing

MsgBox "Transaction Updated!", vbInformation + vbOKOnly, "DAS Version 4.0"
Exit Sub
errh:
MsgBox "Select a Valid File!", vbInformation, "DAS Version 4.0"
Exit Sub
Else
    MsgBox "Invoice not Selected!", vbInformation, "DAS Version 4.0"
End If
    '-----------------------------------------------------
Else
    MsgBox "File Selection Cancelled!", vbInformation + vbOKOnly, "DAS Version 4.0"
    Command7.SetFocus
End If

Else
   MsgBox "Invoice not Selected!", vbInformation, "DAS Version 4.0"
End If
End Sub


Private Sub CYnlist_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSC.Text = CYnlist.Text
    CYnlist.Visible = False
    If MSC.Row <> MSC.Rows - 1 Then MSC.Row = MSC.Row + 1
    MSC.SetFocus
End If
End Sub

Private Sub fileimpcmd_Click()
Dim appXL As Object
Dim appX As Object
Dim TYP As String

If Len(Trim(txtVch.Text)) = 0 Then
    MsgBox "Invoice is not selected!", vbInformation, "DAS Version 4.0"
    txtVch.SetFocus
    Exit Sub
End If

Set appXL = CreateObject("Excel.Application")
CommonDialog1.FileName = ""
CommonDialog1.filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
CommonDialog1.DialogTitle = "Select Mobile Numbers File"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen

CommonDialog1.CancelError = False
If Not Trim(CommonDialog1.FileName) = "" Then
    If LCase(Right(Trim(CommonDialog1.FileName), 3)) = "xls" Then ' Edited
         appXL.Workbooks.Open (CommonDialog1.FileName)  'GetObject(Trim(CommonDialog1.FileName))
    Else
        MsgBox "Selected File is not a Valid SpreadSheet File"
        Exit Sub
    End If
    
    TYP = appXL.Worksheets(1).Cells(1, 1)
    
    If TYP = "Purchase" Then

        Dim RSTEMP As New ADODB.Recordset
        Dim rstemp1 As New ADODB.Recordset
        Dim rstemp2 As New ADODB.Recordset
        Dim pno1 As Integer
    
        On Error GoTo errh
        Dim mn() As String
        Dim mn1() As String
    
        Dim chk_flg As Boolean
        chk_flg = False
    
        Dim pno As Integer
        Dim row_num As Integer
        
        con.Execute "delete from rp2", qc
        
        If RSTEMP.State = 1 Then RSTEMP.Close
        RSTEMP.Open "select chr_one,chr_two from rp2", con, adOpenKeyset, adLockPessimistic
        
        row_num = 3
        tot_num = 0
        
        pno = 0
        While Len(Trim(appXL.Worksheets(1).Cells(row_num, 2))) > 0
            If Len(Trim(Val(appXL.Worksheets(1).Cells(row_num, 2)))) = 10 And Trim(appXL.Worksheets(1).Cells(row_num, 4)) <> "" Then
                If RSTEMP.State = 1 Then RSTEMP.Close
                RSTEMP.Open "select chr_one,chr_two from rp2 WHERE CHR_ONE='" & Trim(appXL.Worksheets(1).Cells(row_num, 2)) & "'", con, adOpenDynamic, adLockPessimistic
                If RSTEMP.EOF Then
                    If rstemp2.State = 1 Then rstemp2.Close
                    rstemp2.Open "select * from avt WHERE mbl_idy='" & Trim(appXL.Worksheets(1).Cells(row_num, 2)) & "'", con, adOpenDynamic, adLockPessimistic
                        If rstemp2.EOF Then
                            
                            RSTEMP.AddNew
                            RSTEMP(0) = Trim(appXL.Worksheets(1).Cells(row_num, 2))
                            RSTEMP(1) = Trim(appXL.Worksheets(1).Cells(row_num, 4))
                            RSTEMP.Update
                            
                        End If
                 End If
            tot_num = Val(tot_num) + 1
            End If
            row_num = Val(row_num) + 1
        Wend
        
        For i = 1 To appXL.Workbooks.count
        appXL.Workbooks(i).Close
        Next
        Set appXL = Nothing
        
        If rstemp1.State = 1 Then rstemp1.Close
        rstemp1.Open "select chr_one,chr_two from RP1", con, adOpenKeyset, adLockPessimistic

        If rstemp1.State = 1 Then rstemp1.Close
        rstemp1.Open "select DISTINCT(CHR_TWO) from rp2", con, adOpenKeyset, adLockPessimistic
        If rstemp1.EOF = False Then

            While rstemp1.EOF = False
            If RSTEMP.State = 1 Then RSTEMP.Close
            RSTEMP.Open "select chr_one,chr_two from rp2 WHERE CHR_TWO='" & Trim(rstemp1(0)) & "'", con, adOpenKeyset, adLockPessimistic
            For K = 1 To MS.Rows
                If SUP_IDY(K) = Trim(rstemp1(0)) Then
                If RSTEMP.EOF = True Then Exit For
                MS.TextMatrix(K, 3) = Trim(RSTEMP(0))
                RSTEMP.MoveNext
                End If
            Next
            rstemp1.MoveNext
            Wend
        End If
        Command3.SetFocus
    con.Execute "delete from rp2", qc
    Exit Sub

End If
    MsgBox "Invalid File!", vbCritical, "DAS Version 4.0"
    appXL.Workbooks.Close
    Set appXL = Nothing
    Exit Sub
Else
    MsgBox "Mobile Numbers File Selection Cancelled!"
    Exit Sub
End If
errh:
    'MsgBox err.Description

End Sub

Private Sub Form_Activate()
'txtStart.Enabled = False ' MODIFIED BY PKB
SSTab1.Tab = 0
txtVch.SetFocus
SSTab1.TabEnabled(3) = False
'SSTab1.TabEnabled(4) = False
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then Shell "c:\windows\calc.exe"
If K > 111 And K < 120 Then If TypeOf M_FRM_MBL.ActiveControl Is TextBox Then M_FRM_MBL.ActiveControl.Text = Trim(FUNKEY(K))
If list1.Visible = False Then
    If K = 117 Then
        DF6
    ElseIf K = 78 And s = 2 Then
    ElseIf K = 83 And s = 2 Then
'        updateavt
    ElseIf K = 81 And s = 2 Then
        cancel
    ElseIf K = 88 And s = 2 Then
        exitfrm
    ElseIf K = 67 And s = 2 Then
    ElseIf K = 68 And s = 2 Then
    ElseIf K = 69 And s = 2 Then
        If M_FRM_MBL.ActiveControl.name = "txtVch" Then invlist
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And s = 2 Then
    ElseIf K = 82 And s = 2 Then
        'MDIForm1.RRReport1.ReportName = App.Path & "\Salesmen list.rrw"
        'MDIForm1.RRReport1.Destination = 1
        'MDIForm1.RRReport1.Printer = "?"
        'MDIForm1.RRReport1.RunReport 1
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If M_FRM_MBL.ActiveControl.name = "sstab1" Then
        If MsgBox("Exit Telecom Interface ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
            Unload Me
        End If
    Else
        If SSTab1.Tab = 3 Then
            If Label49.Caption = "Select Retailer" Then
                SSTab1.Tab = 4
                SSTab1.TabEnabled(3) = False
                txtRet.Enabled = True
                txtRet.SetFocus
            ElseIf Label49 = "Select Invoice" Then
                SSTab1.Tab = 0
                SSTab1.TabEnabled(3) = False
                txtVch.Enabled = True
                txtVch.SetFocus
            ElseIf Label49.Caption = "Select An Invoice" Then
                SSTab1.Tab = 1
                SSTab1.TabEnabled(3) = False
                txtSalInv.Enabled = True
                txtSalInv.SetFocus
            ElseIf Label49 = "Select Invoice (Salesman)" Then
                MSA.CLEAR
                SSTab1.Tab = 2
                SSTab1.TabEnabled(3) = False
                chkAAll.Value = 0
                Option1(4).Value = 1
                txtSalesman = ""
                txtSalesman.Enabled = True
                txtSalesman.SetFocus
            ElseIf Label49 = "Select Party Account" Then
                MSA.CLEAR
                SSTab1.Tab = 2
                SSTab1.TabEnabled(3) = False
                chkAAll.Value = 0
                Option1(3).Value = True
                txtRetailer.Enabled = True
                txtRetailer.SetFocus
            ElseIf Label49 = "Select Mobile No" Then
                MSA.CLEAR
                SSTab1.Tab = 2
                SSTab1.TabEnabled(3) = False
                chkAAll.Value = 0
                Option1(0).Value = True
                txtMobile.Enabled = True
                txtMobile.SetFocus
            ElseIf Label49 = "Select Invoice for Mobile(s) Activation" Then
                MSA.CLEAR
                SSTab1.Tab = 2
                SSTab1.TabEnabled(3) = False
                chkAAll.Value = 0
                Option1(1).Value = True
                txtInvoice.Enabled = True
                txtInvoice.SetFocus
            ElseIf Label49 = "Select Party" Then
                SSTab1.Tab = 2
                SSTab1.TabEnabled(3) = False
                txtRetailer.Enabled = True
                txtRetailer.SetFocus
            ElseIf Label49 = "Select Invoice for Activation Entry" Then
                SSTab1.Tab = 2
                SSTab1.TabEnabled(3) = False
                txtInvoice.Enabled = True
                txtInvoice.SetFocus
            ElseIf Label49 = "Select Mobile Number" Then
                SSTab1.Tab = 2
                SSTab1.TabEnabled(3) = False
                txtMobile.Enabled = True
                txtMobile.SetFocus
            ElseIf Label49 = "Select Invoice (Salesman) for CAF" Then
                SSTab1.Tab = 5
                chkCAll.Value = 0
                chkCAll.Enabled = False
                SSTab1.TabEnabled(3) = False
                txtcSalesman.Enabled = True
                txtcSalesman.SetFocus
                optSmn.Value = True
            ElseIf Label49 = "Select Party Account for CAF" Then
                SSTab1.Tab = 5
                chkCAll.Value = 0
                chkCAll.Enabled = False
                SSTab1.TabEnabled(3) = False
                txtcRetailer.Enabled = True
                optRet.Value = True
                txtcRetailer.SetFocus
            ElseIf Label49 = "Select Mobile NO for CAF" Then
                SSTab1.Tab = 5
                chkCAll.Value = 0
                chkCAll.Enabled = False
                SSTab1.TabEnabled(3) = False
                txtcMobile.Enabled = True
                optMobile.Value = True
                txtcMobile.SetFocus
            ElseIf Label49 = "Select Retailer for CAF" Then
                SSTab1.Tab = 5
                chkCAll.Value = 0
                chkCAll.Enabled = False
                SSTab1.TabEnabled(3) = False
                txtcRetailer.Enabled = True
                txtcRetailer.SetFocus
            ElseIf Label49 = "Select Invoice for Mobile(s) Activation for CAF" Then
                SSTab1.Tab = 5
                chkCAll.Value = 0
                chkCAll.Enabled = False
                SSTab1.TabEnabled(3) = False
                optInvoice.Value = True
                txtcInvoice.Enabled = True
                txtcInvoice.SetFocus
            ElseIf Label49 = "Select Mobile Number for CAF Entry" Then
                SSTab1.Tab = 5
                SSTab1.TabEnabled(3) = False
                txtcMobile.Enabled = True
                txtcMobile.SetFocus
            ElseIf Label49 = "Select Invoice (Salesman) for Churn" Then
                SSTab1.Tab = 6
                SSTab1.TabEnabled(3) = False
                'chkRAll.Value = 0
                optrSmn.Value = True
                txtrSalesman.Enabled = True
                txtrSalesman.SetFocus
            ElseIf Label49 = "Select Retailer for Churn Entry" Then
                SSTab1.Tab = 6
                SSTab1.TabEnabled(3) = False
                optrRet.Value = True
                txtrRetailer.Enabled = True
                txtrRetailer.SetFocus
            ElseIf Label49 = "Select Invoice for Mobile(s) Activation for Churn" Then
                SSTab1.Tab = 6
                SSTab1.TabEnabled(3) = False
'                chkRAll.Value = 0
                optrInvoice.Value = True
                txtrInvoice.Enabled = True
                txtrInvoice.SetFocus
            ElseIf Label49 = "Select Party Account for Churn" Then
                SSTab1.Tab = 6
                SSTab1.TabEnabled(3) = False
                chkRAll.Value = 0
                optrRet.Value = True
                txtrRetailer.Enabled = True
                txtrRetailer.SetFocus
             ElseIf Label49 = "Select Mobile NO for Churn" Then
                SSTab1.Tab = 6
                SSTab1.TabEnabled(3) = False
                chkRAll.Value = 0
                optrMobile.Value = True
                txtrMobile.Enabled = True
                txtrMobile.SetFocus
            ElseIf Label49 = "Select Mobile Number for Churn Entry" Then
                SSTab1.Tab = 6
                SSTab1.TabEnabled(3) = False
                txtrMobile.Enabled = True
                txtrMobile.SetFocus
            ElseIf Label49 = "Select Retailer for Retail Returns" Then
                SSTab1.Tab = 7
                SSTab1.TabEnabled(3) = False
                lblno(0).Visible = False
                lblno(3).Visible = False
                lblno(2).Visible = False
                lblno(1).Visible = False
                txtModRet = ""
                txtModRetName = ""
                txtModRet.SetFocus
            ElseIf Label49 = "Select Invoice for Retail Returns" Then
                SSTab1.Tab = 7
                SSTab1.TabEnabled(3) = False
                txtModInv = ""
                lblno(0).Visible = False
                lblno(3).Visible = False
                lblno(2).Visible = False
                lblno(1).Visible = False
                txtModInv.SetFocus
            ElseIf Label49 = "Select Invoice (Retail Circulation)" Then
                SSTab1.Tab = 7
                SSTab1.TabEnabled(3) = False
                txtModInv = ""
                lblno(0).Visible = False
                lblno(3).Visible = False
                lblno(2).Visible = False
                lblno(1).Visible = False
                txtModInv.SetFocus
            ElseIf M_FRM_MBL.ActiveControl.name = "LVP" Then
                SSTab1.Tab = 0
                SSTab1.TabEnabled(3) = False
                txtVch.SetFocus
                SSTab1.TabEnabled(3) = False
            End If
        ElseIf SSTab1.Tab = 0 Then
            If MS.Rows > 1 Then
                If MsgBox("Save Transactions?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                    Command3_Click
                Else
                    txtVch = ""
                    txtStart = ""
                    txtSup = ""
                    txtVch.SetFocus
                    MS.Rows = 1
                End If
            ElseIf M_FRM_MBL.ActiveControl.name = "txtVch" Then
                Unload Me
            Else
                txtVch.SetFocus
            End If
        ElseIf SSTab1.Tab = 1 Then
            If lv.ListItems.count > 0 Then
                    If MsgBox("Save Transactions?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                        cmdsUpdate_Click
                    Else
                        lv.ListItems.CLEAR
                        txtSalInv = ""
                        txtPartyname = ""
                        txtSmn = ""
                        txtMarket = ""
                        txtSMbl = ""
                        txtSalInv.SetFocus
                    End If
            ElseIf M_FRM_MBL.ActiveControl.name = "LV" Then
                If lv.ListItems.count > 0 Then
                    If MsgBox("Save Transactions?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                        cmdsUpdate_Click
                    Else
                        lv.ListItems.CLEAR
                        txtSalInv = ""
                        txtPartyname = ""
                        txtSmn = ""
                        txtMarket = ""
                        txtSMbl = ""
                        txtSalInv.SetFocus
                    End If
                End If
            ElseIf M_FRM_MBL.ActiveControl.name = "txtSalInv" Then
                SSTab1.Tab = 0
                txtVch.SetFocus
                SSTab1.TabEnabled(3) = False
            End If
        ElseIf SSTab1.Tab = 2 Then
            If Me.ActiveControl.name = "MSA" Then
                If MSA.Rows > 1 Then
                    If Len(Trim(MSA.TextMatrix(1, 1))) > 0 Then
                        If MsgBox("Save Transactions ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                            CMDAUPDATE_Click
                        Else
                            MSA.Rows = 1
                            MSA.Rows = 1
                            txtMobile = ""
                            txtInvoice = ""
                            txtSalesman = ""
                            txtRetailer = ""
                            Ynlist.Visible = False
                            'actDate.SetFocus
                            'Option1(2).Value = True
                            If chkAAll.Enabled = False Then chkAAll.Enabled = True
                            chkAAll.SetFocus
                            chkAAll.Value = 1
                        End If
                    End If
                End If
            ElseIf M_FRM_MBL.ActiveControl.name = "txtMobile" Or M_FRM_MBL.ActiveControl.name = "txtInvoice" Or M_FRM_MBL.ActiveControl.name = "txtRetailer" Or M_FRM_MBL.ActiveControl.name = "txtSalesman" Or M_FRM_MBL.ActiveControl.name = "chkAAll" Then
            
                SSTab1.Tab = 1
                txtSalInv.SetFocus
                SSTab1.TabEnabled(3) = False
            ElseIf M_FRM_MBL.ActiveControl.name = "MSA" Then
                If MSA.Rows > 1 Then
                    If MsgBox("Save Transactions ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                        CMDAUPDATE_Click
                    Else
                        MSA.Rows = 1
                        txtMobile = ""
                        txtInvoice = ""
                        txtSalesman = ""
                        txtRetailer = ""
                        Ynlist.Visible = False
                        'actDate.SetFocus
                        'Option1(2).Value = True
                        chkAAll.SetFocus
                        chkAAll.Value = 1
                    End If
                End If
            Else
                If MS.Rows < 1 Then
                SSTab1.Tab = 1
                txtSalInv.SetFocus
                SSTab1.TabEnabled(3) = False
                End If
            End If
        ElseIf SSTab1.Tab = 4 Then
            If Me.ActiveControl.name = "cmbType" Then
                If cmbType.ListIndex = 0 Then
                    If LVF.ListItems.count > 0 Then
                        If LVR.ListItems.count > 0 Then
                            Command6_Click
                        Else
                            cmbType.SetFocus
                            LVF.ListItems.CLEAR
                            LVR.ListItems.CLEAR
                            If Text2.Visible = True Then Text2.Visible = False
                            txtRet = ""
                            txtRetName = ""
                        End If
                    End If
                Else
                    If MSPC.Rows > 1 Then
                        MSPC.Rows = 1
                        If Text2.Visible = True Then Text2.Visible = False
                        cmbType.SetFocus
                    End If
                End If
                SSTab1.Tab = 2
                'Option1(2).Value = True
                If Text2.Visible = True Then Text2.Visible = False
                'Option1(2).SetFocus
                chkAAll.SetFocus
                chkAAll.Value = 1
            Else
                If cmbType.ListIndex = 0 Then
                    If LVF.ListItems.count > 0 Then
                        If LVR.ListItems.count > 0 Then
                            Command6_Click
                        Else
                            cmbType.SetFocus
                            LVF.ListItems.CLEAR
                            LVR.ListItems.CLEAR
                            txtRet = ""
                            txtRetName = ""
                            If Text2.Visible = True Then Text2.Visible = False
                            txtRet.Enabled = True
                        End If
                    Else
                        cmbType.SetFocus
                        LVF.ListItems.CLEAR
                        LVR.ListItems.CLEAR
                        txtRet = ""
                        txtRetName = ""
                        If Text2.Visible = True Then Text2.Visible = False
                        txtRet.Enabled = True
                    End If
                Else
                    If MSPC.Rows > 1 Then
                        MSPC.Rows = 1
                        If Text2.Visible = True Then Text2.Visible = False
                        cmbType.SetFocus
                    Else
                        If Text2.Visible = True Then Text2.Visible = False
                        cmbType.Enabled = True
                        cmbType.SetFocus
                    End If
                End If
            End If
        ElseIf SSTab1.Tab = 5 Then
            If MSC.Rows > 2 Then
                If MsgBox("Save Transactions?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                    cmdCUpdate_Click
                Else
                    MSC.Rows = 1
                    txtcMobile = ""
                    txtcInvoice = ""
                    txtcSalesman = ""
                    txtcRetailer = ""
                    CYnlist.Visible = False
                    'cafdate.SetFocus
                    'optDate.Value = True
                    If chkCAll.Enabled = False Then chkCAll.Enabled = True
                    chkCAll.SetFocus
                    chkCAll.Value = 1
                End If
            ElseIf M_FRM_MBL.ActiveControl.name = "txtcMobile" Or M_FRM_MBL.ActiveControl.name = "txtcInvoice" Or M_FRM_MBL.ActiveControl.name = "txtcRetailer" Or M_FRM_MBL.ActiveControl.name = "txtcSalesman" Or M_FRM_MBL.ActiveControl.name = "chkCAll" Then
                SSTab1.Tab = 2
                'Option1(2).SetFocus
                If chkAAll.Enabled = False Then chkAAll.Enabled = True
                chkAAll.SetFocus
                chkAAll.Value = 1
                SSTab1.TabEnabled(3) = False
            ElseIf M_FRM_MBL.ActiveControl.name = "MSC" Then
                If MSC.Rows > 1 Then
                    If Len(Trim(MSC.TextMatrix(1, 1))) > 0 Then
                        If MsgBox("Save Transactions?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                            cmdCUpdate_Click
                        Else
                            MSC.Rows = 1
                            txtcMobile = ""
                            txtcInvoice = ""
                            txtcSalesman = ""
                            txtcRetailer = ""
                            CYnlist.Visible = False
                            'cafdate.SetFocus
                            'optDate.Value = True
                            If chkCAll.Enabled = False Then chkCAll.SetFocus
                            chkCAll.SetFocus
                            chkCAll.Value = 1
                        End If
                    End If
                End If
            End If
        ElseIf SSTab1.Tab = 6 Then
            If M_FRM_MBL.ActiveControl.name = "txtrMobile" Or M_FRM_MBL.ActiveControl.name = "txtrInvoice" Or M_FRM_MBL.ActiveControl.name = "txtrRetailer" Or M_FRM_MBL.ActiveControl.name = "txtrSalesman" Or M_FRM_MBL.ActiveControl.name = "chkRAll" Then
                SSTab1.Tab = 5
                'cafdate.SetFocus
                'optDate.Value = True
                If chkCAll.Enabled = False Then chkCAll.SetFocus
                chkCAll.SetFocus
                chkCAll.Value = 1
                SSTab1.TabEnabled(3) = False
            ElseIf MSR.Rows > 1 Then
                If Len(Trim(MSR.TextMatrix(1, 1))) > 0 Then
                    If MsgBox("Save Transactions?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                        cmdRUpdate_Click
                    Else
                        MSR.Rows = 1
                        txtrInvoice = ""
                        txtrMobile = ""
                        txtrSalesman = ""
                        txtrRetailer = ""
                        RYnlist.Visible = False
                        'ChrDate.SetFocus
                        'optrDate.Value = True
                        If chkRAll.Enabled = False Then chkRAll.Enabled = True
                        chkRAll.SetFocus
                        chkRAll.Value = 1
                    End If
                End If
            ElseIf M_FRM_MBL.ActiveControl.name = "MSR" Then
                If MSR.Rows > 1 Then
                    If MsgBox("Save Transactions?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                        cmdRUpdate_Click
                    Else
                        MSR.Rows = 1
                        txtrInvoice = ""
                        txtrMobile = ""
                        txtrSalesman = ""
                        txtrRetailer = ""
                        RYnlist.Visible = False
                        'ChrDate.SetFocus
                        'optrDate.Value = True
                        If chkRAll.Enabled = False Then chkRAll.Enabled = True
                        chkRAll.SetFocus
                        chkRAll.Value = 1
                    End If
                End If
            End If
        ElseIf SSTab1.Tab = 7 Then
            If Me.ActiveControl.name = "modDate" Then
                LVMOD.ListItems.CLEAR
                txtModRet = ""
                txtModRetName = ""
                txtModInv = ""
                SSTab1.Tab = 6
            Else
                If LVMOD.ListItems.count >= 1 Then
                    If MsgBox("Update Transactions?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                        cmdModUpdate_Click
                    Else
                        lblno(0).Visible = False
                        lblno(3).Visible = False
                        lblno(2).Visible = False
                        lblno(1).Visible = False
                        txtModRet = ""
                        txtModRetName = ""
                        txtModInv = ""
                        LVMOD.ListItems.CLEAR
                        modDate.SetFocus
                    End If
                Else
                    txtModRet = ""
                    txtModRetName = ""
                    txtModInv = ""
                    modDate.SetFocus
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub Form_Load()
SETYEAR Me
With MS
    .Cols = 9
    .Rows = 1
    .TextMatrix(0, 0) = "Invoice ID"
    .ColWidth(0) = 0
    .TextMatrix(0, 1) = "S.No"
    .ColWidth(1) = 600
    .TextMatrix(0, 2) = "SKU Name"
    .ColWidth(2) = 2555
    .TextMatrix(0, 3) = "Mobile Number"
    .ColAlignment(3) = 7
    .ColWidth(3) = 1900
    .ColAlignment(3) = 7
    .TextMatrix(0, 4) = "Invoice Date"
    .ColWidth(4) = 0
    .ColWidth(5) = 0
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .ColWidth(8) = 0
End With
End Sub

Private Sub Form_Unload(cancel As Integer)
If RSINV.State = 1 Then RSINV.Close
Set RSINV = Nothing
If rschk.State = 1 Then rschk.Close
Set rschk = Nothing
If rsAvt.State = 1 Then rsAvt.Close
Set rsAvt = Nothing
If RSINL.State = 1 Then RSINL.Close
Set RSINL = Nothing
If rsprd.State = 1 Then rsprd.Close
Set rsprd = Nothing
If rsact.State = 1 Then rsact.Close
Set rsact = Nothing
If RSSMN.State = 1 Then RSSMN.Close
Set RSSMN = Nothing
If rsmar.State = 1 Then rsmar.Close
Set rsmar = Nothing
On Error Resume Next
If rsimp.State = 1 Then rsimp.Close
Set rsimp = Nothing
If rsEdt.State = 1 Then rsEdt.Close
Set rsEdt = Nothing
If rsEdt1.State = 1 Then rsEdt1.Close
Set rsEdt1 = Nothing
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim lv_chk As Boolean
lv_chk = list_chk(Item.Index)

Dim temp As String
Dim temp1 As String
Dim temp2 As String
Dim temp3 As String
Dim temp4 As String
Dim temp5 As String
Dim temp6 As String
Dim temp_chk As Integer

 If Item.Checked = True Then
    
    If lv_chk = False Then
    Item.Checked = False
    MsgBox "Cannot Select More Than Invoiced for this SKU Starter Pack! ", vbCritical, "DAS Version 4.0"
    lv.SetFocus
    Exit Sub
    End If
    
    For i = 1 To lv.ListItems.count
    If Trim(lv.ListItems(i).Text) = Trim(Item.Text) Then
    temp_chk = 0
    
    For i1 = 1 To lv.ListItems.count
    If lv.ListItems(i1).Checked = False Then
        temp_chk = i1
        Exit For
    End If
    Next
    
    If temp_chk = 0 Then
        Exit Sub
    End If
    
        If i < temp_chk Then
        Exit For
        Else
        temp = lv.ListItems(i)
        temp1 = lv.ListItems(i).SubItems(1)
        temp2 = lv.ListItems(i).SubItems(2)
        temp3 = lv.ListItems(i).SubItems(3)
        temp4 = lv.ListItems(i).SubItems(4)
        temp5 = lv.ListItems(i).SubItems(5)
'        temp6 = LV.ListItems(I).SubItems(6)
    
        For K = i To temp_chk + 1 Step -1
        If K >= 2 Then
        lv.ListItems(K).Text = lv.ListItems(K - 1).Text
        lv.ListItems(K).SubItems(1) = lv.ListItems(K - 1).SubItems(1)
        lv.ListItems(K).SubItems(2) = lv.ListItems(K - 1).SubItems(2)
        lv.ListItems(K).SubItems(3) = lv.ListItems(K - 1).SubItems(3)
        lv.ListItems(K).SubItems(4) = lv.ListItems(K - 1).SubItems(4)
        lv.ListItems(K).SubItems(5) = lv.ListItems(K - 1).SubItems(5)
'        lv.ListItems(K).SubItems(6) = lv.ListItems(K - 1).SubItems(6)
        End If
        Next
    
        If temp_chk > 0 Then
            lv.ListItems(temp_chk).Text = temp
            lv.ListItems(temp_chk).SubItems(1) = temp1
            lv.ListItems(temp_chk).SubItems(2) = temp2
            lv.ListItems(temp_chk).SubItems(3) = temp3
            lv.ListItems(temp_chk).SubItems(4) = temp4
            lv.ListItems(temp_chk).SubItems(5) = temp5
'            lv.ListItems(temp_chk).SubItems(6) = temp6
            lv.ListItems(temp_chk).Checked = True
            lv.ListItems(i).Checked = False
        End If
            CHK1 = 1
        End If
        Exit For
    End If
    Next
Else
'Dim I1 As Integer
i1 = Item.Index

If Item.Index = lv.ListItems.count Then Exit Sub

temp = lv.ListItems(i1).Text
temp1 = lv.ListItems(i1).SubItems(1)
temp2 = lv.ListItems(i1).SubItems(2)
temp3 = lv.ListItems(i1).SubItems(3)
temp4 = lv.ListItems(i1).SubItems(4)
temp5 = lv.ListItems(i1).SubItems(5)
'temp6 = LV.ListItems(I1).SubItems(6)

For K = i1 To lv.ListItems.count - 1
    lv.ListItems(K).Text = lv.ListItems(K + 1).Text
    lv.ListItems(K).SubItems(1) = lv.ListItems(K + 1).SubItems(1)
    lv.ListItems(K).SubItems(2) = lv.ListItems(K + 1).SubItems(2)
    lv.ListItems(K).SubItems(3) = lv.ListItems(K + 1).SubItems(3)
    lv.ListItems(K).SubItems(4) = lv.ListItems(K + 1).SubItems(4)
    lv.ListItems(K).SubItems(5) = lv.ListItems(K + 1).SubItems(5)
    If lv.ListItems(K + 1).Checked = True Then
        lv.ListItems(K).Checked = True
        lv.ListItems(K + 1).Checked = False
    End If
'    LV.ListItems(K).SubItems(6) = LV.ListItems(K + 1).SubItems(6)
Next
i1 = lv.ListItems.count
 lv.ListItems(i1).Text = temp
 lv.ListItems(i1).SubItems(1) = temp1
 lv.ListItems(i1).SubItems(2) = temp2
 lv.ListItems(i1).SubItems(3) = temp3
 lv.ListItems(i1).SubItems(4) = temp4
 lv.ListItems(i1).SubItems(5) = temp5

End If

End Sub


Private Sub LVMOD_ItemCheck(ByVal Item As MSComctlLib.ListItem)

Dim CHK_RET_FLG As Boolean
CHK_RET_FLG = False

If Item.Checked = False Then

    If Trim(Item.SubItems(2)) = "Yes" Or Trim(Item.SubItems(2)) = "SAL" Then
        If Trim(Item.SubItems(2)) = "SAL" Then
            MsgBox "Cannot Modify Sale Transaction!", vbInformation + vbOKOnly, "DAS Version 4.0"
        Else
            MsgBox "Cannot Modify Activated Transaction!", vbInformation + vbOKOnly, "DAS Version 4.0"
        End If
        Item.Checked = True
        LVMOD.SetFocus
    End If
    
End If

End Sub

Private Sub LVP_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If LVP.SortOrder = lvwAscending Then
LVP.SortOrder = lvwDescending
ElseIf LVP.SortOrder = lvwDescending Then
LVP.SortOrder = lvwAscending
End If
LVP.SortKey = CH.Index - 1
End Sub

Private Sub LVP_DblClick()
LVP_KeyDown 13, 0
End Sub

Private Sub LVP_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ls As ListItem

If KeyCode = 13 Then
    If LVP.ListItems.count > 0 Then
        If Trim(Label49.Caption) = "Select Invoice for Adjust Returns" Then
            txtModInv = LVP.SelectedItem
                        
            If Mid(Trim(txtModInv), 1, 3) = "PUR" Or Mid(Trim(txtModInv), 1, 3) = "OPS" Or Mid(Trim(txtModInv), 1, 3) = "STI" Or Mid(Trim(txtModInv), 1, 3) = "SRN" Then
            Dim chk_rst As New ADODB.Recordset
            Dim chk_rst1 As New ADODB.Recordset
            Dim chk_prd As New ADODB.Recordset
                LVSRN.Visible = False
                modDate.MinDate = "01/01/1999"
                modDate.MaxDate = "12/31/2015"
                modDate = LVP.SelectedItem.ListSubItems(1)
                If rsdir.State = 1 Then rsdir.Close
                rsdir.Open "SELECT DIR_ALS,DIR_MSG FROM DIR WHERE DIR_IDY='CUR' AND DIR_IDX='FYR'", con, adOpenStatic
                If rsAvt.State = 1 Then rsAvt.Close
                Set rsAvt = Nothing
                
'                rsAvt.Open "select inv_idy,inv_dat,spl_idy,prd_idy,txn_typ,mbl_idy from avt where inv_idy='" & Trim(txtModInv) & "' and inv_dat=ctod('" & invdat.Value & "') GROUP BY PRD_IDY", CON, adOpenKeyset, adlockpessimistic
        
                LVMOD.ColumnHeaders.CLEAR
                LVMOD.ColumnHeaders.Add , , "Mobile No.", 1500
                LVMOD.ColumnHeaders.Add , , "Product Name", 2500
                LVMOD.ColumnHeaders.Add , , "SAL. Status", 1200
                LVMOD.ListItems.CLEAR
                
        chk_rst.Open "select count(*),prd_idy from avt where INV_idy='" & Trim(txtModInv) & "' group by prd_idy", con, adOpenStatic
        
        If chk_rst.EOF = False Then
            While chk_rst.EOF = False
            If chk_rst1.State = 1 Then chk_rst1.Close
            chk_rst1.Open "select prd_idy,prd_qty from inv where inv_idy='" & Trim(txtModInv) & "' and prd_idy='" & chk_rst(1) & "'", con, adOpenStatic
            If Val(chk_rst(0)) <> Val(chk_rst1(1)) Then
                cg = True
                If chk_prd.State = 1 Then chk_prd.Close
                chk_prd.Open "select PRD_IDY,prd_nme from prd where prd_idy='" & chk_rst(1) & "' and prd_idy='" & chk_rst(1) & "'", con, adOpenStatic
                On Error GoTo aaaa
                PRDID(retno).Caption = Trim(chk_prd(0))
                lblno(retno).Caption = Trim(chk_prd(1)) & " : " & (chk_rst1(1))
                TT.Caption = Val(retno)
                lblno(retno).Visible = True
                If rsaMod.State = 1 Then rsaMod.Close
                rsaMod.Open "select mbl_idy,SAL_TYP from avt where INV_idy='" & Trim(txtModInv) & "' AND PRD_IDY='" & chk_prd(0) & "'", con, adOpenStatic
                If Not rsaMod.EOF Then
                While Not rsaMod.EOF
                    Set ls = LVMOD.ListItems.Add(, , Trim(rsaMod(0)) & "")
                    ls.ListSubItems.Add , , Trim(chk_prd(1)) & ""
                    ls.ListSubItems.Add , , Trim(rsaMod(1)) & ""
                    ls.Checked = True
                    rsaMod.MoveNext
                Wend
            End If
aaaa:
            End If
            retno = retno + 1
            chk_rst.MoveNext
            Wend
        End If
        
        If chk_rst.State = 1 Then chk_rst.Close
        If chk_rst1.State = 1 Then chk_rst1.Close
'        If chk_prd.State = 1 Then chk_prd.Close
        
        If cg = True Then
           
            If rsaMod.State = 1 Then rsaMod.Close
            rsaMod.Open "select mbl_idy,avt_sts from avt where INV_idy='" & Trim(txtModInv) & "'", con, adOpenStatic
            If Not rsaMod.EOF Then
'                txtModInv = IVID
''                LVMOD.ColumnHeaders.Clear
''                LVMOD.ColumnHeaders.Add , , "Mobile No.", 2500
''                LVMOD.ColumnHeaders.Add , , "Activation Status", 1900
''                LVMOD.ListItems.Clear
'                While Not rsaMod.EOF
'                    Set ls = LVMOD.ListItems.Add(, , Trim(rsaMod(0)) & "")
'                    ls.ListSubItems.Add , , Trim(rsaMod(1)) & ""
'                    ls.Checked = True
'                    rsaMod.MoveNext
'                Wend
                SSTab1.Tab = 7
                LVMOD.SetFocus
            Else
                MsgBox "Transactions Does not Exist for the selected Invoice!", vbInformation + vbOKOnly, "DAS Version 4.0"
                SSTab1.Tab = 7
                txtModInv.Enabled = True
                txtModInv.SetFocus
            End If
          Else
         
          SSTab1.Tab = 7
          SSTab1.TabEnabled(3) = False
          txtModInv = ""
          MsgBox "Products not Returned!", vbInformation + vbOKOnly, "DAS Version 4.0"
          txtModInv.SetFocus
        End If
                
'                If rsact.State = 1 Then rsact.Close
'                rsact.Open "select act_idy,act_nme from act", CON, adOpenKeyset, adlockpessimistic
'                If rsprd.State = 1 Then rsprd.Close
'                rsprd.Open "select prd_idy,prd_nme,GRP_IDY,ALT_IDY from prd", CON, adOpenKeyset, adlockpessimistic
'                MS.Rows = rsAvt.RecordCount + 1
'                ReDim SUP_IDY(MS.Rows) As String
'                r = 1
'                If rsAvt.RecordCount > 0 Then
'                    mbid = False
'
'                    SU1 = ""
'                    CHK_SUP = False
'                    Dim scode As Integer
'                    scode = 0
'                    While Not rsAvt.EOF
'                        If rsact.State = 1 Then rsact.Close
'                        rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & rsAvt("spl_idy") & "'", CON, adOpenKeyset, adlockpessimistic
'                        If Not rsact.EOF Then txtSup = Trim(rsact(1))
'                        If rsprd.State = 1 Then rsprd.Close
'                        rsprd.Open "select prd_idy,prd_nme,GRP_IDY,ALT_IDY from prd WHERE prd_idy='" & rsAvt("prd_idy") & "'", CON, adOpenKeyset, adlockpessimistic
'                        MS.TextMatrix(r, 1) = Val(r)
'                        If Not rsprd.EOF Then MS.TextMatrix(r, 2) = Trim(rsprd(1))
'                        If Not rsprd.EOF Then MS.TextMatrix(r, 8) = Trim(rsprd(2))
'
'                        'CHECKING THE SUPPLIER Id OF PRODUCT
'                        If IsNull(rsprd(3)) = False And Trim(rsprd(3)) <> "" Then
'                            SUP_IDY(r) = Trim(rsprd(3))
'                        Else
'                            CHK_SUP = True
'                        End If
'                        MS.TextMatrix(r, 3) = Val(Trim(rsAvt(5)) & "")
'                        If Val(Trim(rsAvt(5)) & "") > 0 Then mbid = True
'                        MS.TextMatrix(r, 5) = Trim(rsprd(0)) & ""
'                        MS.TextMatrix(r, 6) = Trim(rsAvt(4)) & ""
'                        If Not rsact.EOF Then MS.TextMatrix(r, 7) = Trim(rsact(0)) & ""
'                        r = r + 1
'                        rsAvt.MoveNext
'                    Wend
'                    Command1.Enabled = True
'                    cmdGen.Enabled = True
'                End If
'                If CHK_SUP = False Then
'                    SSTab1.Tab = 0
'                    SSTab1.TabEnabled(3) = False
'
'                    If mbid = True Then
'                        MS.Col = 4
'                        MS.SetFocus
'                    Else
'                        txtStart.SetFocus
'                    End If
'                Else
'                    SSTab1.Tab = 0
'                    SSTab1.TabEnabled(3) = False
'                    cancel
'                    MsgBox "Supplier's Product Code not Specified to the Product!", vbCritical, "DAS Version 4.0"
'                    txtVch.SetFocus
'                End If
                
'            ElseIf Mid(Trim(txtVch), 1, 3) = "OPS" Then
'
'                LVSRN.Visible = False
'                invdat = LVP.SelectedItem.ListSubItems(1)
'                If rsAvt.State = 1 Then rsAvt.Close
'                Set rsAvt = Nothing
'                rsAvt.Open "select inv_idy,inv_dat,spl_idy,prd_idy,txn_typ,mbl_idy from avt where inv_idy='" & Trim(txtVch) & "'", CON, adOpenKeyset, adlockpessimistic
'                If rsact.State = 1 Then rsact.Close
'                rsact.Open "select act_idy,act_nme from act", CON, adOpenKeyset, adlockpessimistic
'                If rsprd.State = 1 Then rsprd.Close
'                rsprd.Open "select prd_idy,prd_nme,GRP_IDY from prd", CON, adOpenKeyset, adlockpessimistic
'                MS.Rows = rsAvt.RecordCount + 1
'                ReDim SUP_IDY(MS.Rows) As String
'                r = 1
'                SU1 = 0
'                If rsAvt.RecordCount > 0 Then
'                    mbid = False
'                    While Not rsAvt.EOF
'                        If rsact.State = 1 Then rsact.Close
'                        rsact.Open "select act_idy,act_nme from ACT WHERE ACT_idy='" & rsAvt("spl_idy") & "'", CON, adOpenKeyset, adlockpessimistic
'                        If Not rsact.EOF Then txtSup = Trim(rsact(1))
'                        If rsprd.State = 1 Then rsprd.Close
'                        rsprd.Open "select prd_idy,prd_nme,GRP_IDY,ALT_IDY from prd WHERE prd_idy='" & rsAvt("prd_idy") & "'", CON, adOpenKeyset, adlockpessimistic
'                        MS.TextMatrix(r, 1) = Val(r)
'                        If Not rsprd.EOF Then MS.TextMatrix(r, 2) = Trim(rsprd(1))
'                        If Not rsprd.EOF Then MS.TextMatrix(r, 8) = Trim(rsprd(2))
'
'                        If IsNull(rsprd(3)) = False And Trim(rsprd(3)) <> "" Then
'                            SUP_IDY(r) = Trim(rsprd(3))
'                        Else
'                            CHK_SUP = True
'                        End If
'
'                        MS.TextMatrix(r, 3) = Val(Trim(rsAvt(5)) & "")
'                        If Val(Trim(rsAvt(5)) & "") > 0 Then mbid = True
'                        MS.TextMatrix(r, 5) = Trim(rsprd(0)) & ""
'                        MS.TextMatrix(r, 6) = Trim(rsAvt(4)) & ""
'                        If Not rsact.EOF Then MS.TextMatrix(r, 7) = Trim(rsact(0)) & ""
'                        r = r + 1
'                        rsAvt.MoveNext
'                      Wend
'                    Command1.Enabled = True
'                    cmdGen.Enabled = True
'                End If
'                If CHK_SUP = False Then
'                SSTab1.Tab = 0
'                SSTab1.TabEnabled(3) = False
'                If mbid = True Then
'                    MS.Col = 3
'                    MS.SetFocus
'                Else
'                    txtStart.SetFocus
'                End If
'                Else
'                    SSTab1.Tab = 0
'                    SSTab1.TabEnabled(3) = False
'                    cancel
'                    MsgBox "Supplier's Product Code not Specified to the Product!", vbCritical, "DAS Version 4.0"
'                    txtVch.SetFocus
'                End If
'            Else
'                LVSRN.Visible = True
'                MS.Visible = False
'                If RSINL.State = 1 Then RSINL.Close
'                Set RSINL = Nothing
'                RSINL.Open "select cst_idy,act_nme from inl,act where inl.cst_idy like act.act_idy and inv_idy='" & Trim(txtVch) & "'", CON, adOpenKeyset, adlockpessimistic
'                If Not RSINL.EOF Then
'                    txtSup = Trim(RSINL(1))
'                    If rsAvt.State = 1 Then rsAvt.Close
'                    Set rsAvt = Nothing
'                    rsAvt.Open "select mbl_idy,sal_dat from avt where cst_idy='" & Trim(RSINL(0)) & "' and sal_typ='SAL'", CON, adOpenKeyset, adlockpessimistic
'                    LVSRN.ColumnHeaders.Clear
'                    LVSRN.ColumnHeaders.Add , , "Mobile Number", 2500
'                    LVSRN.ColumnHeaders.Add , , "Invoice Date", 1500
'                    LVSRN.ListItems.Clear
'                    While Not rsAvt.EOF
'                        Set ls = LVSRN.ListItems.Add(, , rsAvt(0))
'                        ls.ListSubItems.Add , , rsAvt(1)
'                        rsAvt.MoveNext
'                    Wend
'                    SSTab1.Tab = 0
'                    SSTab1.TabEnabled(3) = False
'                    LVSRN.SetFocus
'                End If
            End If
        ElseIf Trim(Label49.Caption) = "Select Invoice" Then    ' Pur,Ops,Sti,Srn Invoices List
            txtVch = LVP.SelectedItem
            If Mid(Trim(txtVch), 1, 3) = "PUR" Or Mid(Trim(txtVch), 1, 3) = "STI" Or Mid(Trim(txtVch), 1, 3) = "SRN" Then
                LVSRN.Visible = False
                invdat.MinDate = "01/01/1999"
                invdat.MaxDate = "12/31/2015"
                invdat = LVP.SelectedItem.ListSubItems(1)
                If rsdir.State = 1 Then rsdir.Close
                rsdir.Open "SELECT DIR_ALS,DIR_MSG FROM DIR WHERE DIR_IDY='CUR' AND DIR_IDX='FYR'", con, adOpenStatic
                If rsAvt.State = 1 Then rsAvt.Close
                Set rsAvt = Nothing
                'rsAvt.Open "select inv_idy,inv_dat,spl_idy,prd_idy,txn_typ,mbl_idy from avt where inv_idy='" & Trim(txtVch) & "' and inv_dat>=ctod('" & Format(CDate(rsdir("dir_als")), "mm/DD/yyyy") & "') and inv_dat<=ctod('" & Format(CDate(rsdir("dir_msg")), "mm/DD/yyyy") & "') and (mbl_idy is null or mbl_idy='')", CON, adOpenKeyset, adlockpessimistic
                rsAvt.Open "select inv_idy,inv_dat,spl_idy,prd_idy,txn_typ,mbl_idy from avt where inv_idy='" & Trim(txtVch) & "' and inv_dat=ctod('" & invdat.Value & "') and (mbl_idy is null or mbl_idy='')", con, adOpenKeyset, adLockPessimistic
                
        '       If rsinv.State = 1 Then rsinv.Close
        '       rsinv.Open "select sum(prd_qty) from inv where inv_idy='" & Trim(txtVch) & "'", CON, adOpenKeyset, adlockpessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_idy,prd_nme,GRP_IDY,ALT_IDY from prd", con, adOpenKeyset, adLockPessimistic
                MS.Rows = rsAvt.RecordCount + 1
                ReDim SUP_IDY(MS.Rows) As String
                r = 1
                If rsAvt.RecordCount > 0 Then
                    mbid = False
                    
                    SU1 = ""
                    CHK_SUP = False
'                    Dim scode As Integer
                    scode = 0
                    While Not rsAvt.EOF
                        If rsact.State = 1 Then rsact.Close
                        rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & rsAvt("spl_idy") & "'", con, adOpenKeyset, adLockPessimistic
                        If Not rsact.EOF Then txtSup = Trim(rsact(1))
                        If rsprd.State = 1 Then rsprd.Close
                        rsprd.Open "select prd_idy,prd_nme,GRP_IDY,ALT_IDY from prd WHERE prd_idy='" & rsAvt("prd_idy") & "'", con, adOpenKeyset, adLockPessimistic
                        MS.TextMatrix(r, 1) = Val(r)
                        If Not rsprd.EOF Then MS.TextMatrix(r, 2) = Trim(rsprd(1))
                        If Not rsprd.EOF Then MS.TextMatrix(r, 8) = Trim(rsprd(2))
                        
                        'CHECKING THE SUPPLIER Id OF PRODUCT
                        If IsNull(rsprd(3)) = False And Trim(rsprd(3)) <> "" Then
                            SUP_IDY(r) = Trim(rsprd(3))
                        Else
                            CHK_SUP = True
                        End If
                        MS.TextMatrix(r, 3) = Val(Trim(rsAvt(5)) & "")
                        If Val(Trim(rsAvt(5)) & "") > 0 Then mbid = True
                        MS.TextMatrix(r, 5) = Trim(rsprd(0)) & ""
                        MS.TextMatrix(r, 6) = Trim(rsAvt(4)) & ""
                        If Not rsact.EOF Then MS.TextMatrix(r, 7) = Trim(rsact(0)) & ""
                        r = r + 1
                        rsAvt.MoveNext
                    Wend
                    Command1.Enabled = True
                    cmdGen.Enabled = True
                End If
                If CHK_SUP = False Then
                    SSTab1.Tab = 0
                    SSTab1.TabEnabled(3) = False
                    
                    If mbid = True Then
                        MS.Col = 4
                        MS.SetFocus
                    Else
                        txtStart.SetFocus
                    End If
                Else
                    SSTab1.Tab = 0
                    SSTab1.TabEnabled(3) = False
                    cancel
                    MsgBox "Supplier's Product Code not Specified to the Product!", vbCritical, "DAS Version 4.0"
                    txtVch.SetFocus
                End If
                
            ElseIf Mid(Trim(txtVch), 1, 3) = "OPS" Then
                
                LVSRN.Visible = False
                invdat = LVP.SelectedItem.ListSubItems(1)
                If rsAvt.State = 1 Then rsAvt.Close
                Set rsAvt = Nothing
                rsAvt.Open "select inv_idy,inv_dat,spl_idy,prd_idy,txn_typ,mbl_idy from avt where inv_idy='" & Trim(txtVch) & "'", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_idy,prd_nme,GRP_IDY from prd", con, adOpenKeyset, adLockPessimistic
                MS.Rows = rsAvt.RecordCount + 1
                ReDim SUP_IDY(MS.Rows) As String
                r = 1
                SU1 = 0
                If rsAvt.RecordCount > 0 Then
                    mbid = False
                    While Not rsAvt.EOF
                        If rsact.State = 1 Then rsact.Close
                        rsact.Open "select act_idy,act_nme from ACT WHERE ACT_idy='" & rsAvt("spl_idy") & "'", con, adOpenKeyset, adLockPessimistic
                        If Not rsact.EOF Then txtSup = Trim(rsact(1))
                        If rsprd.State = 1 Then rsprd.Close
                        rsprd.Open "select prd_idy,prd_nme,GRP_IDY,ALT_IDY from prd WHERE prd_idy='" & rsAvt("prd_idy") & "'", con, adOpenKeyset, adLockPessimistic
                        MS.TextMatrix(r, 1) = Val(r)
                        If Not rsprd.EOF Then MS.TextMatrix(r, 2) = Trim(rsprd(1))
                        If Not rsprd.EOF Then MS.TextMatrix(r, 8) = Trim(rsprd(2))
                        
                        If IsNull(rsprd(3)) = False And Trim(rsprd(3)) <> "" Then
                            SUP_IDY(r) = Trim(rsprd(3))
                        Else
                            CHK_SUP = True
                        End If
                        
                        MS.TextMatrix(r, 3) = Val(Trim(rsAvt(5)) & "")
                        If Val(Trim(rsAvt(5)) & "") > 0 Then mbid = True
                        MS.TextMatrix(r, 5) = Trim(rsprd(0)) & ""
                        MS.TextMatrix(r, 6) = Trim(rsAvt(4)) & ""
                        If Not rsact.EOF Then MS.TextMatrix(r, 7) = Trim(rsact(0)) & ""
                        r = r + 1
                        rsAvt.MoveNext
                      Wend
                    Command1.Enabled = True
                    cmdGen.Enabled = True
                End If
                If CHK_SUP = False Then
                SSTab1.Tab = 0
                SSTab1.TabEnabled(3) = False
                If mbid = True Then
                    MS.Col = 3
                    MS.SetFocus
                Else
                    txtStart.SetFocus
                End If
                Else
                    SSTab1.Tab = 0
                    SSTab1.TabEnabled(3) = False
                    cancel
                    MsgBox "Supplier's Product Code not Specified to the Product!", vbCritical, "DAS Version 4.0"
                    txtVch.SetFocus
                End If
            Else
                LVSRN.Visible = True
                MS.Visible = False
                If RSINL.State = 1 Then RSINL.Close
                Set RSINL = Nothing
                RSINL.Open "select cst_idy,act_nme from inl,act where inl.cst_idy like act.act_idy and inv_idy='" & Trim(txtVch) & "'", con, adOpenKeyset, adLockPessimistic
                If Not RSINL.EOF Then
                    txtSup = Trim(RSINL(1))
                    If rsAvt.State = 1 Then rsAvt.Close
                    Set rsAvt = Nothing
                    rsAvt.Open "select mbl_idy,sal_dat from avt where cst_idy='" & Trim(RSINL(0)) & "' and sal_typ='SAL'", con, adOpenKeyset, adLockPessimistic
                    LVSRN.ColumnHeaders.CLEAR
                    LVSRN.ColumnHeaders.Add , , "Mobile Number", 2500
                    LVSRN.ColumnHeaders.Add , , "Invoice Date", 1500
                    LVSRN.ListItems.CLEAR
                    While Not rsAvt.EOF
                        Set ls = LVSRN.ListItems.Add(, , rsAvt(0))
                        ls.ListSubItems.Add , , rsAvt(1)
                        rsAvt.MoveNext
                    Wend
                    SSTab1.Tab = 0
                    SSTab1.TabEnabled(3) = False
                    LVSRN.SetFocus
                End If
            End If
        ElseIf Trim(Label49.Caption) = "Select An Invoice" Then   ' Sal,Sto,Prn Invoices List
        
            Dim pct_rst As New ADODB.Recordset
            Dim pct_rst1 As New ADODB.Recordset
            Dim PRD_CNT As Integer
            txtSalInv = LVP.SelectedItem
            SQTY = Val(LVP.SelectedItem.ListSubItems(2).Text)
            sinvdat.MinDate = "01/01/1999"
            sinvdat.MaxDate = "12/31/2015"
            sinvdat = LVP.SelectedItem.ListSubItems(3).Text
            TTYP = LVP.SelectedItem.ListSubItems(5).Text
            PID = LVP.SelectedItem.ListSubItems(8).Text
            SID = LVP.SelectedItem.ListSubItems(9).Text
            PrID = LVP.SelectedItem.ListSubItems(7).Text
            
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            'rsAvt.Open "SELECT MBL_IDY,PRD_IDY,CST_IDY,INV_IDY FROM AVT WHERE PRD_IDY IN(SELECT PRD_IDY FROM INV WHERE INV_IDY='" & Trim(txtSalInv) & "') and mbl_idy<>' ' and sal_typ is null or sal_typ='SRN' or sal_typ<>'SAL' order by mbl_idy", CON, adOpenKeyset, adlockpessimistic
            rsAvt.Open "SELECT MBL_IDY,PRD_IDY,CST_IDY,INV_IDY FROM AVT WHERE PRD_IDY IN(SELECT PRD_IDY FROM INV WHERE inv_dat=ctod('" & sinvdat & "') and INV_IDY='" & Trim(txtSalInv) & "') and mbl_idy<>' ' and sal_typ is null or sal_typ='SRN' or sal_typ<>'SAL' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            
            If pct_rst.State = 1 Then pct_rst.Close
            pct_rst.Open "seleCT prd.PRD_nme,SUM(inv.prd_qty),INV.PRD_IDY FROM INV,prd WHERE PRD.EXC_INF='Y' AND INV.inv_dat=ctod('" & sinvdat & "') and prd.prd_idy=inv.prd_idy and TRIM(INV_IDY)='" & Trim(txtSalInv) & "' GROUP BY PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
            
            ppp = 0
            
            While pct_rst.EOF = False
            
                If pct_rst1.State = 1 Then pct_rst1.Close
                pct_rst1.Open "select count(prd_idy) from avt where sal_dat=ctod('" & sinvdat.Value & "') and sal_idy='" & Trim(txtSalInv) & "' and prd_idy='" & Trim(pct_rst(2)) & "'", con, adOpenKeyset, adLockPessimistic
                PRD_CNT = 0
                
                On Error GoTo sw
                If pct_rst1.EOF = False Then
                PRD_CNT = pct_rst(1) - pct_rst1(0)
                Else
                PRD_CNT = pct_rst(1)
                End If
                Ptno(ppp) = Trim(PRD_CNT) & " : " & Trim(pct_rst(0))
                Ptno(ppp).Visible = True
                pn.Caption = ppp
                pn.Visible = False
sw:
                pct_rst.MoveNext
                ppp = ppp + 1
            
            Wend
            
            pct_rst.Close
            Set pct_rst = Nothing
            
            lv.ColumnHeaders.CLEAR
            lv.ListItems.CLEAR
            lv.ColumnHeaders.Add , , "Mobile ID", 1800
            lv.ColumnHeaders.Add , , "SKU Name", 2500
            lv.ColumnHeaders.Add , , "Salesman", 0
            lv.ColumnHeaders.Add , , "Invoice ID", 0
            lv.ColumnHeaders.Add , , "ID", 0
            lv.ColumnHeaders.Add , , "Pr Id", 0
            lv.ColumnHeaders.Add , , "smn Id", 0
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_idy,prd_nme from prd", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
            If RSSMN.State = 1 Then RSSMN.Close
            RSSMN.Open "select smn_idy,smn_nme from smn", con, adOpenKeyset, adLockPessimistic
            If RSSMN.RecordCount > 0 Then RSSMN.MoveFirst
            RSSMN.Find "smn_idy='" & SID & "'"
            If Not RSSMN.EOF Then txtSmn = RSSMN(1)
            If rsmar.State = 1 Then rsmar.Close
            rsmar.Open "SELECT mar.mar_idy,mar_nme,act_idy from mar,act where mar.mar_idy like act.mar_idy and act_idy='" & Trim(PID) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsmar.EOF Then txtMarket = rsmar(1)
            Dim SNO As Integer
            SNO = 1
            If rsAvt.RecordCount > 0 Then
                Do While Not rsAvt.EOF
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_idy,prd_nme from PRD WHERE PRD_idy='" & Trim(rsAvt(1)) & "'", con, adOpenKeyset, adLockPessimistic
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(PID) & "'", con, adOpenKeyset, adLockPessimistic
                    Set lst = lv.ListItems.Add(, , rsAvt(0))
                    If Not rsprd.EOF Then lst.ListSubItems.Add , , Trim(rsprd(1)) & ""
                    If RSSMN.RecordCount > 0 Then
'                    RSSMN.MoveFirst
'                    RSSMN.Find "smn_idy='" & sid & "'"
                    lst.ListSubItems.Add , , RSSMN(1)
                    Else
                    lst.ListSubItems.Add , , " "
                    End If
                    lst.ListSubItems.Add , , rsAvt(3)
                    lst.ListSubItems.Add , , Trim(PID)
                    lst.ListSubItems.Add , , Trim(rsprd(0)) & ""
                    If Not rsprd.EOF Then
                    lst.ListSubItems.Add , , RSSMN(0) & ""
                    Else
                    lst.ListSubItems.Add , , " " & ""
                    End If
                    rsAvt.MoveNext
                    SNO = SNO + 1
                Loop
                txtPartyname = rsact(1)
                Label4 = SQTY 'RSINV(0)
                Label4.Visible = False
                If rsAvt.EOF = True Then
                    SSTab1.Tab = 1
                    lv.SetFocus
                    SSTab1.TabEnabled(3) = False
                End If
            Else
                MsgBox "Mobile Numbers not Found for the selected Invoice Product!", vbOKOnly, "DAS Version 4.0"
                SSTab1.Tab = 1
                txtSalInv = ""
                txtSmn = ""
                txtMarket = ""
                txtSMbl = ""
                txtSalInv.SetFocus
                SSTab1.TabEnabled(3) = False
            End If
        ElseIf Trim(Label49.Caption) = "Select Invoice for Mobile(s) Activation" Then
            invid = LVP.SelectedItem
            txtInvoice = Trim(invid)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,sal_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' and (avt_sts is null or avt_sts='No' or avt_sts='') and sal_idy='" & invid & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSA.Rows = 1
                MSA.Cols = 7
                MSA.TextMatrix(0, 0) = "SNO"
                MSA.ColWidth(0) = 700
                MSA.TextMatrix(0, 1) = "Mobile ID"
                MSA.ColWidth(1) = 1300
                MSA.ColAlignment(1) = 1
                MSA.TextMatrix(0, 2) = "Invoice ID"
                MSA.ColWidth(2) = 1300
                MSA.TextMatrix(0, 3) = "Status"
                MSA.ColWidth(3) = 800
                MSA.TextMatrix(0, 4) = "SKU Name"
                MSA.ColWidth(4) = 2200
                MSA.TextMatrix(0, 5) = "Party"
                MSA.ColWidth(5) = 1740
                MSA.TextMatrix(0, 6) = "Invoice Date"
                MSA.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSA.Rows = rsAvt.RecordCount + 1
                MSA1.Rows = rsAvt.RecordCount + 1
                r = 1
                MSA1.Cols = 6
                While Not rsAvt.EOF
                    MSA.TextMatrix(r, 0) = r
                    MSA.TextMatrix(r, 1) = Trim(rsAvt(0))
                    MSA.TextMatrix(r, 2) = Trim(rsAvt(1))
                    MSA.TextMatrix(r, 3) = "No"
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSA.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSA.TextMatrix(r, 5) = Trim(rsact(1))
                    MSA.TextMatrix(r, 6) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                r = 1
                rsAvt.MoveFirst
                While Not rsAvt.EOF
                    MSA1.TextMatrix(r, 0) = Trim(rsAvt(0))
                    MSA1.TextMatrix(r, 1) = Trim(rsAvt(1))
                    MSA1.TextMatrix(r, 2) = "No"
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSA1.TextMatrix(r, 3) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSA1.TextMatrix(r, 4) = Trim(rsact(1))
                    MSA1.TextMatrix(r, 5) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 2
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                avtdate.SetFocus
            End If
        ElseIf Trim(Label49.Caption) = "Select Mobile No" Then
        MBLID = LVP.SelectedItem.Text
        If RSINL.State = 1 Then RSINL.Close
        Set RSINL = Nothing
        RSINL.Open "select mbl_idy,sal_idy,sal_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' AND (avt_sts is null or avt_sts='No' or avt_sts='') and mbl_idy='" & Trim(MBLID) & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
        If RSINL.RecordCount > 0 Then
            
            txtMobile.Text = MBLID
    '       MSA.FixedRows = 0
            MSA.Rows = 1
            MSA.Cols = 7
            MSA.TextMatrix(0, 0) = "SNO"
             MSA.ColWidth(0) = 700
            MSA.TextMatrix(0, 1) = "Mobile ID"
            MSA.ColWidth(1) = 1300
            MSA.ColAlignment(1) = 1
            MSA.TextMatrix(0, 2) = "Invoice ID"
            MSA.ColWidth(2) = 1300
            MSA.TextMatrix(0, 3) = "Status"
            MSA.ColWidth(3) = 800
            MSA.TextMatrix(0, 4) = "SKU Name"
            MSA.ColWidth(4) = 2200
            MSA.TextMatrix(0, 5) = "Party"
            MSA.ColWidth(5) = 1740
            MSA.TextMatrix(0, 6) = "Invoice Date"
            MSA.ColWidth(6) = 1200
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
            MSA.Rows = RSINL.RecordCount + 1
    '        MSA.FixedRows = 1
            r = 1
            While Not RSINL.EOF
                MSA.TextMatrix(r, 0) = r
                MSA.TextMatrix(r, 1) = Trim(RSINL(0))
                If IsNull(RSINL(1)) = True Then
                MSA.TextMatrix(r, 2) = " "
                Else
                MSA.TextMatrix(r, 2) = Trim(RSINL(1))
                End If
                MSA.TextMatrix(r, 3) = "No" 'Trim(RSINL(6)) & ""
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsprd.EOF Then MSA.TextMatrix(r, 4) = Trim(rsprd(0))
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsact.EOF Then MSA.TextMatrix(r, 5) = Trim(rsact(1))
                If IsNull(RSINL(2)) = True Then
                MSA.TextMatrix(r, 6) = " "
                Else
                MSA.TextMatrix(r, 6) = Trim(RSINL(2))
                End If
                r = r + 1
                RSINL.MoveNext
            Wend
            SSTab1.Tab = 2
            SSTab1.TabEnabled(3) = False
            avtdate.SetFocus
            'MSA.Col = 2
            'MSA.SetFocus
        Else
            MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
            SSTab1.Tab = 2
            'actDate.SetFocus
            'Option1(2).Value = True
            chkAAll.SetFocus
            chkAAll.Value = 1
            SSTab1.TabEnabled(3) = False
        End If
        ElseIf Trim(Label49.Caption) = "Select Party Account" Then
            retid = LVP.SelectedItem.ListSubItems(1).Text
            If rsact.State = 1 Then rsact.Close
            rsact.Open "SELECT ACT_IDY,ACT_NME FROM ACT WHER ACT_IDY='" & Trim(retid) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then txtRetailer = rsact(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,sal_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' and (avt_sts is null or avt_sts='No' or avt_sts='') and cst_idy='" & retid & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSA.Rows = 1
                MSA.Cols = 7
                MSA.TextMatrix(0, 0) = "Mobile ID"
                MSA.ColWidth(0) = 700
                MSA.TextMatrix(0, 1) = "Mobile ID"
                MSA.ColWidth(1) = 1300
                MSA.ColAlignment(1) = 1
                MSA.TextMatrix(0, 2) = "Invoice ID"
                MSA.ColWidth(2) = 1300
                MSA.TextMatrix(0, 3) = "Status"
                MSA.ColWidth(3) = 800
                MSA.TextMatrix(0, 4) = "SKU Name"
                MSA.ColWidth(4) = 2200
                MSA.TextMatrix(0, 5) = "Party"
                MSA.ColWidth(5) = 1740
                MSA.TextMatrix(0, 6) = "Invoice Date"
                MSA.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSA.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSA.TextMatrix(r, 0) = r
                    MSA.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSA.TextMatrix(r, 2) = " "
                    Else
                    MSA.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSA.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSA.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSA.TextMatrix(r, 5) = Trim(rsact(1))
                    If IsNull(rsAvt(2)) = True Then
                    MSA.TextMatrix(r, 6) = " "
                    Else
                    MSA.TextMatrix(r, 6) = Trim(rsAvt(2))
                    End If
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 2
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                avtdate.SetFocus
                'MSA.Col = 2
                'MSA.SetFocus
            End If
        ElseIf Trim(Label49.Caption) = "Select Invoice (Salesman)" Then
        
        'MODIFIED BY PKB
        
        If rsAvt.State = 1 Then rsAvt.Close
        txtSalesman = LVP.SelectedItem.Text
        smnid = LVP.SelectedItem.ListSubItems(1).Text
        rsAvt.Open "select SMN_idy from aVt WHERE SMN_IDY = '" & Trim(smnid) & "'", con, adOpenKeyset, adLockPessimistic
        If rsAvt.RecordCount > 0 Then           'MODIFIED BY PKB
        
            smnid = ""
            smnid = LVP.SelectedItem.ListSubItems(1).Text
            If RSSMN.State = 1 Then RSSMN.Close
            RSSMN.Open "SELECT SMN_IDY,SMN_NME FROM SMN WHERE SMN_IDY='" & Trim(smnid) & "'", con, adOpenKeyset, adLockPessimistic
            If Not RSSMN.EOF Then txtSalesman = RSSMN(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,sal_dat,prd_idy,cst_idy,txn_typ,avt_sts,smn_idy from avt where sal_typ='SAL' and (avt_sts is null or avt_sts='No' or avt_sts='') and smn_idy='" & Trim(smnid) & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSA.Rows = 1
                MSA.Cols = 7
                MSA.TextMatrix(0, 0) = "SNO"
                MSA.ColWidth(0) = 700
                MSA.TextMatrix(0, 1) = "Mobile ID"
                MSA.ColWidth(1) = 1300
                MSA.ColAlignment(1) = 1
                MSA.TextMatrix(0, 2) = "Invoice ID"
                MSA.ColWidth(2) = 1300
                MSA.TextMatrix(0, 3) = "Status"
                MSA.ColWidth(3) = 800
                MSA.TextMatrix(0, 4) = "SKU Name"
                MSA.ColWidth(4) = 2200
                MSA.TextMatrix(0, 5) = "Party"
                MSA.ColWidth(5) = 1740
                MSA.TextMatrix(0, 6) = "Invoice Date"
                MSA.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSA.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSA.TextMatrix(r, 0) = r
                    MSA.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSA.TextMatrix(r, 2) = " "
                    Else
                    MSA.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSA.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSA.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSA.TextMatrix(r, 5) = Trim(rsact(1))
                    If IsNull(rsAvt(2)) = True Then
                    MSA.TextMatrix(r, 6) = " "
                    Else
                    MSA.TextMatrix(r, 6) = Trim(rsAvt(2))
                    End If
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 2
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                MSA.Col = 3
                MSA.SetFocus
                avtdate.SetFocus
            
            End If
            Else                        'MODIFIED BY PKB
            SSTab1.Tab = 2
            MsgBox "No records found", vbInformation + vbOKOnly, "DAS Version 4.0"
            txtSalesman = ""
            If chkAAll.Enabled = False Then chkAAll.Enabled = True
            chkAAll.SetFocus
            'LVP.SetFocus       'MODIFIED BY PKB
            End If               'MODIFIED BY PK'B
        ElseIf Trim(Label49.Caption) = "Select Invoice for Mobile(s) Activation for CAF" Then
            invid = LVP.SelectedItem
            txtcInvoice = Trim(invid)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' and avt_sts='Yes' and (caf_sts is null or caf_sts<>'Yes' or caf_sts='') and sal_idy='" & invid & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSC.Rows = 1
                MSC.Cols = 7
                MSC.TextMatrix(0, 0) = "SNO"
                MSC.ColWidth(0) = 700
                MSC.TextMatrix(0, 1) = "Mobile ID"
                MSC.ColWidth(1) = 1300
                MSC.ColAlignment(1) = 1
                MSC.TextMatrix(0, 2) = "Invoice ID"
                MSC.ColWidth(2) = 1300
                MSC.TextMatrix(0, 3) = "Status"
                MSC.ColWidth(3) = 800
                MSC.TextMatrix(0, 4) = "SKU Name"
                MSC.ColWidth(4) = 2200
                MSC.TextMatrix(0, 5) = "Party"
                MSC.ColWidth(5) = 1740
                MSC.TextMatrix(0, 6) = "Act. Date"
                MSC.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSC.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSC.TextMatrix(r, 0) = r
                    MSC.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSC.TextMatrix(r, 2) = " "
                    Else
                    MSC.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSC.TextMatrix(r, 3) = "No"
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSC.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSC.TextMatrix(r, 5) = Trim(rsact(1))
                    MSC.TextMatrix(r, 6) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 5
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                MSC.Col = 3
                MSC.SetFocus
            End If
        ElseIf Trim(Label49.Caption) = "Select Mobile NO for CAF" Then
        
        MBLID = LVP.SelectedItem.Text
        If RSINL.State = 1 Then RSINL.Close
        Set RSINL = Nothing
        RSINL.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts,caf_sts from avt where sal_typ='SAL' AND avt_sts='Yes' and (caf_sts is null or caf_sts<>'Yes' or caf_sts='') and mbl_idy='" & Trim(MBLID) & "'", con, adOpenKeyset, adLockPessimistic
        If RSINL.RecordCount > 0 Then
            MSC.Rows = 1
            MSC.Cols = 7
            MSC.TextMatrix(0, 0) = "SNO"
            MSC.ColWidth(0) = 700
            MSC.TextMatrix(0, 1) = "Mobile ID"
            MSC.ColWidth(1) = 1300
            MSC.ColAlignment(1) = 1
            MSC.TextMatrix(0, 2) = "Invoice ID"
            MSC.ColWidth(2) = 1300
            MSC.TextMatrix(0, 3) = "Status"
            MSC.ColWidth(3) = 800
            MSC.TextMatrix(0, 4) = "SKU Name"
            MSC.ColWidth(4) = 2200
            MSC.TextMatrix(0, 5) = "Party"
            MSC.ColWidth(5) = 1740
            MSC.TextMatrix(0, 6) = "Act. Date"
            MSC.ColWidth(6) = 1200

            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
            MSC.Rows = RSINL.RecordCount + 1
            r = 1
            While Not RSINL.EOF
                MSC.TextMatrix(r, 0) = r
                MSC.TextMatrix(r, 1) = Trim(RSINL(0))
                If IsNull(RSINL(1)) = True Then
                MSC.TextMatrix(r, 2) = ""
                Else
                MSC.TextMatrix(r, 2) = Trim(RSINL(1))
                End If
                txtcMobile.Text = Trim(RSINL(0))
                MSC.TextMatrix(r, 3) = "No" 'Trim(RSINL(7)) & ""
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsprd.EOF Then MSC.TextMatrix(r, 4) = Trim(rsprd(0))
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsact.EOF Then MSC.TextMatrix(r, 5) = Trim(rsact(1))
                MSC.TextMatrix(r, 6) = Trim(RSINL(2))
                r = r + 1
                RSINL.MoveNext
            Wend
            SSTab1.Tab = 5
            SSTab1.TabEnabled(3) = False
            MSC.Col = 3
            MSC.SetFocus
        Else
            MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
            SSTab1.Tab = 5
            'cafdate.SetFocus
            'optDate.Value = True
            chkCAll.SetFocus
            chkCAll.Value = 1
            SSTab1.TabEnabled(3) = False
        End If
        ElseIf Trim(Label49.Caption) = "Select Party Account for CAF" Then
            retid = LVP.SelectedItem.ListSubItems(1).Text
            If rsact.State = 1 Then rsact.Close
            rsact.Open "SELECT ACT_IDY,ACT_NME FROM ACT WHER ACT_IDY='" & Trim(retid) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then txtcRetailer = rsact(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' and avt_sts='Yes' and (caf_sts is null or caf_sts<>'Yes' or caf_sts='') and cst_idy='" & retid & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSC.Rows = 1
                MSC.Cols = 7
                MSC.TextMatrix(0, 0) = "SNO"
                MSC.ColWidth(0) = 700
                MSC.TextMatrix(0, 1) = "Mobile ID"
                MSC.ColWidth(1) = 1300
                MSC.ColAlignment(1) = 1
                MSC.TextMatrix(0, 2) = "Invoice ID"
                MSC.ColWidth(2) = 1300
                MSC.TextMatrix(0, 3) = "Status"
                MSC.ColWidth(3) = 800
                MSC.TextMatrix(0, 4) = "SKU Name"
                MSC.ColWidth(4) = 2200
                MSC.TextMatrix(0, 5) = "Party"
                MSC.ColWidth(5) = 1740
                MSC.TextMatrix(0, 6) = "Act. Date"
                MSC.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSC.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSC.TextMatrix(r, 0) = r
                    MSC.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSC.TextMatrix(r, 2) = ""
                    Else
                    MSC.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSC.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSC.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSC.TextMatrix(r, 5) = Trim(rsact(1))
                    MSC.TextMatrix(r, 6) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 5
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                MSC.Col = 3
                MSC.SetFocus
            End If
        ElseIf Trim(Label49.Caption) = "Select Invoice (Salesman) for CAF" Then
        
        
        If rsAvt.State = 1 Then rsAvt.Close
        txtcSalesman = LVP.SelectedItem.Text
        smnid = LVP.SelectedItem.ListSubItems(1).Text
        rsAvt.Open "select SMN_idy from aVt WHERE SMN_IDY = '" & Trim(smnid) & "'", con, adOpenKeyset, adLockPessimistic
        If rsAvt.RecordCount > 0 Then           'CONDITION ADDED BY PKB
        
            smnid = ""
            smnid = LVP.SelectedItem.ListSubItems(1).Text
        
       '     smnid = LVP.SelectedItem.ListSubItems(5).Text
            If RSSMN.State = 1 Then RSSMN.Close
            RSSMN.Open "SELECT SMN_IDY,SMN_NME FROM SMN WHERE SMN_IDY='" & Trim(smnid) & "'", con, adOpenKeyset, adLockPessimistic
            If Not RSSMN.EOF Then txtcSalesman = RSSMN(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts,smn_idy from avt where sal_typ='SAL' and avt_sts='Yes' and (caf_sts is null or caf_sts<>'Yes' or caf_sts='') and smn_idy='" & Trim(smnid) & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSC.Rows = 1
                MSC.Cols = 7
                MSC.TextMatrix(0, 0) = "SNO"
                MSC.ColWidth(0) = 1300
                MSC.TextMatrix(0, 1) = "Mobile ID"
                MSC.ColWidth(1) = 1300
                MSC.ColAlignment(1) = 1
                MSC.TextMatrix(0, 2) = "Invoice ID"
                MSC.ColWidth(2) = 1300
                MSC.TextMatrix(0, 3) = "Status"
                MSC.ColWidth(3) = 800
                MSC.TextMatrix(0, 4) = "SKU Name"
                MSC.ColWidth(4) = 2200
                MSC.TextMatrix(0, 5) = "Party"
                MSC.ColWidth(5) = 1740
                MSC.TextMatrix(0, 6) = "Act. Date"
                MSC.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSC.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSC.TextMatrix(r, 0) = r
                    MSC.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSC.TextMatrix(r, 2) = ""
                    Else
                    MSC.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSC.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSC.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSC.TextMatrix(r, 5) = Trim(rsact(1))
                    MSC.TextMatrix(r, 6) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 5
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                MSC.Col = 3
                MSC.SetFocus
                End If
            Else 'MODIFIED BY PK'B
            SSTab1.Tab = 5
            MsgBox "No records found", vbInformation + vbOKOnly, "DAS Version 4.0"
            'LVP.SetFocus
            txtcSalesman = ""
            chkCAll.SetFocus
            End If
        ElseIf Trim(Label49.Caption) = "Select Invoice for Mobile(s) Activation for Churn" Then
            invid = LVP.SelectedItem
            txtrInvoice = Trim(invid)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' and avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') and sal_idy='" & invid & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSR.Rows = 1
                MSR.Cols = 7
                MSR.TextMatrix(0, 0) = "SNO"
                MSR.ColWidth(0) = 700
                MSR.TextMatrix(0, 1) = "Mobile ID"
                MSR.ColWidth(1) = 1300
                MSR.ColAlignment(1) = 1
                MSR.TextMatrix(0, 2) = "Invoice ID"
                MSR.ColWidth(2) = 1300
                MSR.TextMatrix(0, 3) = "Status"
                MSR.ColWidth(3) = 800
                MSR.TextMatrix(0, 4) = "SKU Name"
                MSR.ColWidth(4) = 2200
                MSR.TextMatrix(0, 5) = "Party"
                MSR.ColWidth(5) = 1740
                MSR.TextMatrix(0, 6) = "Act. Date"
                MSR.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSR.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSR.TextMatrix(r, 0) = r
                    MSR.TextMatrix(r, 1) = Trim(rsAvt(0))
                    MSR.TextMatrix(r, 2) = Trim(rsAvt(1))
                    MSR.TextMatrix(r, 3) = "No"
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSR.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSR.TextMatrix(r, 5) = Trim(rsact(1))
                    MSR.TextMatrix(r, 6) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 6
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                avtcdate.SetFocus
                'MSR.Col = 2
                'MSR.SetFocus
            End If
        ElseIf Trim(Label49.Caption) = "Select Mobile NO for Churn" Then
            MBLID = LVP.SelectedItem.Text
            txtrMobile.Text = MBLID
            If RSINL.State = 1 Then RSINL.Close
            Set RSINL = Nothing
                RSINL.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts,caf_sts from avt where sal_typ='SAL' AND avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') and mbl_idy='" & Trim(MBLID) & "'", con, adOpenKeyset, adLockPessimistic
                If RSINL.RecordCount > 0 Then
                    MSR.Rows = 1
                    MSR.Cols = 7
                    MSR.TextMatrix(0, 0) = "SNO"
                    MSR.ColWidth(0) = 700
                    MSR.TextMatrix(0, 1) = "Mobile ID"
                    MSR.ColWidth(1) = 1300
                    MSR.ColAlignment(1) = 1
                    MSR.TextMatrix(0, 2) = "Invoice ID"
                    MSR.ColWidth(2) = 1300
                    MSR.TextMatrix(0, 3) = "Status"
                    MSR.ColWidth(3) = 800
                    MSR.TextMatrix(0, 4) = "SKU Name"
                    MSR.ColWidth(4) = 2200
                    MSR.TextMatrix(0, 5) = "Party"
                    MSR.ColWidth(5) = 1740
                    MSR.TextMatrix(0, 6) = "Act. Date"
                    MSR.ColWidth(6) = 1200
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                    MSR.Rows = RSINL.RecordCount + 1
                    r = 1
                    While Not RSINL.EOF
                        MSR.TextMatrix(r, 0) = r
                        MSR.TextMatrix(r, 1) = Trim(RSINL(0))
                        If IsNull(RSINL(1)) = True Then
                        MSR.TextMatrix(r, 2) = " "
                        Else
                        MSR.TextMatrix(r, 2) = Trim(RSINL(1))
                        End If
                        MSR.TextMatrix(r, 3) = "No" 'Trim(RSINL(6)) & ""
                        If rsprd.State = 1 Then rsprd.Close
                        rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
                        If Not rsprd.EOF Then MSR.TextMatrix(r, 4) = Trim(rsprd(0))
                        If rsact.State = 1 Then rsact.Close
                        rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
                        If Not rsact.EOF Then MSR.TextMatrix(r, 5) = Trim(rsact(1))
                        MSR.TextMatrix(r, 6) = Trim(RSINL(2))
                        r = r + 1
                        RSINL.MoveNext
                    Wend
                    SSTab1.Tab = 6
                    SSTab1.TabEnabled(3) = False
                    avtcdate.SetFocus
                    'MSR.Col = 2
                    'MSR.SetFocus
                Else
                    MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
                    SSTab1.Tab = 6
                    txtrMobile.SetFocus
                    SSTab1.TabEnabled(3) = False
                End If
        ElseIf Trim(Label49.Caption) = "Select Party Account for Churn" Then
            retid = LVP.SelectedItem.ListSubItems(1).Text
            If rsact.State = 1 Then rsact.Close
            rsact.Open "SELECT ACT_IDY,ACT_NME FROM ACT WHER ACT_IDY='" & Trim(retid) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then txtrRetailer = rsact(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' and avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') and cst_idy='" & retid & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSR.Rows = 1
                MSR.Cols = 7
                MSR.TextMatrix(0, 0) = "SNO"
                MSR.ColWidth(0) = 700
                MSR.TextMatrix(0, 1) = "Mobile ID"
                MSR.ColWidth(1) = 1300
                MSR.ColAlignment(1) = 1
                MSR.TextMatrix(0, 2) = "Invoice ID"
                MSR.ColWidth(2) = 1300
                MSR.TextMatrix(0, 3) = "Status"
                MSR.ColWidth(3) = 800
                MSR.TextMatrix(0, 4) = "SKU Name"
                MSR.ColWidth(4) = 2200
                MSR.TextMatrix(0, 5) = "Party"
                MSR.ColWidth(5) = 1740
                MSR.TextMatrix(0, 6) = "Act. Date"
                MSR.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSR.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSR.TextMatrix(r, 0) = r
                    MSR.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSR.TextMatrix(r, 2) = " "
                    Else
                    MSR.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSR.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSR.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSR.TextMatrix(r, 5) = Trim(rsact(1))
                    MSR.TextMatrix(r, 6) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 6
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                avtcdate.SetFocus
            End If
        ElseIf Trim(Label49.Caption) = "Select Invoice (Salesman) for Churn" Then
        
        'MODIFIED BY PKB
         If rsAvt.State = 1 Then rsAvt.Close
            txtrSalesman = LVP.SelectedItem.Text
            smnid = LVP.SelectedItem.ListSubItems(1).Text
            rsAvt.Open "select SMN_idy from aVt WHERE SMN_IDY = '" & Trim(smnid) & "'", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then

            smnid = LVP.SelectedItem.ListSubItems(1).Text   ' INSPITE OF 5 PLACED 1 IN LSI
            If RSSMN.State = 1 Then RSSMN.Close
            RSSMN.Open "SELECT SMN_IDY,SMN_NME FROM SMN WHERE SMN_IDY='" & Trim(smnid) & "'", con, adOpenKeyset, adLockPessimistic
            If Not RSSMN.EOF Then txtrSalesman = RSSMN(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts,smn_idy from avt where sal_typ='SAL' and avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') and smn_idy='" & Trim(smnid) & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSR.Rows = 1
                MSR.Cols = 7
                MSR.TextMatrix(0, 0) = "SNO"
                MSR.ColWidth(0) = 700
                MSR.TextMatrix(0, 1) = "Mobile ID"
                MSR.ColWidth(1) = 1300
                MSR.ColAlignment(1) = 1
                MSR.TextMatrix(0, 2) = "Invoice ID"
                MSR.ColWidth(2) = 1300
                MSR.TextMatrix(0, 3) = "Status"
                MSR.ColWidth(3) = 800
                MSR.TextMatrix(0, 4) = "SKU Name"
                MSR.ColWidth(4) = 2200
                MSR.TextMatrix(0, 5) = "Party"
                MSR.ColWidth(5) = 1740
                MSR.TextMatrix(0, 6) = "Act. Date"
                MSR.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSR.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSR.TextMatrix(r, 0) = r
                    MSR.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSR.TextMatrix(r, 2) = " "
                    Else
                    MSR.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSR.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSR.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSR.TextMatrix(r, 5) = Trim(rsact(1))
                    MSR.TextMatrix(r, 6) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 6
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                avtcdate.SetFocus
                End If
               Else
                SSTab1.Tab = 6
                MsgBox "No records found", vbInformation + vbOKOnly, "DAS Version 4.0"
                txtrSalesman = ""
                chkRAll.SetFocus
            End If
        ElseIf Label49.Caption = "Select Retailer" Then
            'change
                Dim Lst1 As ListSubItem
                Set Lst1 = LVP.SelectedItem.ListSubItems.Item(1)
                If rs.State = 1 Then rs.Close
                Set rs = Nothing
                rs.Open "select * from act where AcT_idy like '" & Trim(Lst1.Text) & "'", con, adOpenStatic
                txtRet = Trim(rs("act_idy"))
                txtRetName = Trim(rs("act_nme"))
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                SSTab1.Tab = 4
                txtRet.Enabled = False
                dispDisMbl
                dispRetMbl (Trim(txtRet))
                If LVF.ListItems.count = 0 Then
                    MsgBox "Sales does not Exist!", vbInformation + vbOKOnly, "DAS Version 4.0"
                    LVR.ListItems.CLEAR
                    txtRet.Enabled = True
                    txtRet = ""
                    txtRetName = ""
                    txtSmn = ""
                    txtMarket = ""
                    txtSMbl = ""
                    cmbType.SetFocus
                    Exit Sub
                End If
            'change
        ElseIf Label49.Caption = "Select Party Account (Retail Circulation)" Then
            Dim lst2 As ListSubItem
            Set lst2 = LVP.SelectedItem.ListSubItems.Item(1)
            txtModRet = lst2
            txtModRetName = LVP.SelectedItem.Text
            SSTab1.Tab = 7
            txtModInv.Enabled = True
            txtModInv.SetFocus
        ElseIf Label49.Caption = "Select Invoice (Retail Circulation)" Then
        
        retno = 0
        cg = False
        IVID = Trim(LVP.SelectedItem.Text)
        txtModInv = IVID
        modDate.Value = LVP.SelectedItem.SubItems(2)
        LVMOD.ColumnHeaders.CLEAR
        LVMOD.ColumnHeaders.Add , , "Mobile No.", 1500
        LVMOD.ColumnHeaders.Add , , "Product Name", 2500
        LVMOD.ColumnHeaders.Add , , "Act. Status", 1200
        LVMOD.ListItems.CLEAR
                
        chk_rst.Open "select count(*),prd_idy from avt where sal_idy='" & Trim(IVID) & "' group by prd_idy", con, adOpenStatic
        
        If chk_rst.EOF = False Then
            While chk_rst.EOF = False
            If chk_rst1.State = 1 Then chk_rst1.Close
            chk_rst1.Open "select prd_idy,prd_qty from inv where inv_idy='" & Trim(IVID) & "' and prd_idy='" & chk_rst(1) & "'", con, adOpenStatic
            If Val(chk_rst(0)) <> Val(chk_rst1(1)) Then
                cg = True
                If chk_prd.State = 1 Then chk_prd.Close
                chk_prd.Open "select PRD_IDY,prd_nme from prd where prd_idy='" & chk_rst(1) & "' and prd_idy='" & chk_rst(1) & "'", con, adOpenStatic
                On Error GoTo aaaa1
                PRDID(retno).Caption = Trim(chk_prd(0))
                lblno(retno).Caption = Trim(chk_prd(1)) & " : " & (chk_rst1(1))
                TT.Caption = Val(retno)
                lblno(retno).Visible = True
                If rsaMod.State = 1 Then rsaMod.Close
                rsaMod.Open "select mbl_idy,avt_sts from avt where sal_idy='" & Trim(IVID) & "' AND PRD_IDY='" & chk_prd(0) & "'", con, adOpenStatic
                If Not rsaMod.EOF Then
                While Not rsaMod.EOF
                    Set ls = LVMOD.ListItems.Add(, , Trim(rsaMod(0)) & "")
                    ls.ListSubItems.Add , , Trim(chk_prd(1)) & ""
                    ls.ListSubItems.Add , , Trim(rsaMod(1)) & ""
                    ls.Checked = True
                    rsaMod.MoveNext
                Wend
            End If
aaaa1:
            End If
            retno = retno + 1
            chk_rst.MoveNext
            Wend
        End If
        
        If chk_rst.State = 1 Then chk_rst.Close
        If chk_rst1.State = 1 Then chk_rst1.Close
        If chk_prd.State = 1 Then chk_prd.Close
        
        If cg = True Then
           
            If rsaMod.State = 1 Then rsaMod.Close
            rsaMod.Open "select mbl_idy,avt_sts from avt where sal_idy='" & Trim(IVID) & "'", con, adOpenStatic
            If Not rsaMod.EOF Then
'                txtModInv = IVID
''                LVMOD.ColumnHeaders.Clear
''                LVMOD.ColumnHeaders.Add , , "Mobile No.", 2500
''                LVMOD.ColumnHeaders.Add , , "Activation Status", 1900
''                LVMOD.ListItems.Clear
'                While Not rsaMod.EOF
'                    Set ls = LVMOD.ListItems.Add(, , Trim(rsaMod(0)) & "")
'                    ls.ListSubItems.Add , , Trim(rsaMod(1)) & ""
'                    ls.Checked = True
'                    rsaMod.MoveNext
'                Wend
                SSTab1.Tab = 7
                LVMOD.SetFocus
            Else
                MsgBox "Transactions Does not Exist for the selected Invoice!", vbInformation + vbOKOnly, "DAS Version 4.0"
                SSTab1.Tab = 7
                txtModInv.Enabled = True
                txtModInv.SetFocus
            End If
          Else
         
          SSTab1.Tab = 7
          SSTab1.TabEnabled(3) = False
          txtModInv = ""
          MsgBox "Products not Returned!", vbInformation + vbOKOnly, "DAS Version 4.0"
          txtModInv.SetFocus
        End If
        End If
    End If
End If

End Sub

Private Sub LVP_KeyPress(K As Integer)
If K = 27 Then
'If Label49 = "Select Party Account for Churn" Then
If SSTab1.Tab = 6 Then
chkRAll.Enabled = False
ElseIf SSTab1.Tab = 2 Then
chkAAll.Enabled = False
ElseIf SSTab1.Tab = 5 Then
chkRAll.Enabled = False
End If

'txtrRetailer.SetFocus
'End If
'End If
''Unload Me
'
'        LVP.Visible = False
'        SSTab1.TabEnabled(3) = False
'        SSTab1.Tab = 5
'        Frame1.Enabled = True
'        If txtcRetailer.Visible = False Then txtcRetailer.Visible = True
'        If txtcRetailer.Enabled = False Then txtcRetailer.Enabled = True
'       ' txtcSalesman.SetFocus
'        chkCAll.Enabled = False
'        optSmn.Value = True
'        'Modified by Parvathi
'        txtcInvoice.Visible = True
'        txtcSalesman.Visible = True
'        txtcRetailer.Visible = True
'        txtcMobile.Visible = True
'
'        optSmn.Enabled = True
'        optInvoice.Enabled = True
'        optRet.Enabled = True
'        optMobile.Enabled = True
'        txtcRetailer.SetFocus
End If
End Sub

Private Sub modDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Option2.SetFocus
'    txtModRet.Enabled = True
'    txtModRet.SetFocus
End If
End Sub

Private Sub MS_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
    If MS.Col = 3 Then
        'If rschk.State = 1 Then rschk.Close
        'rschk.Open "select mbl_idy from avt where sal_typ='SAL' AND mbl_idy='" & Trim(MS.Text) & "'", CON, adOpenKeyset, adlockpessimistic
        'If rschk.EOF Then
            ROWNO = MS.Row
            MS.Text = ""
            txtTemp.Height = MS.CellHeight
            txtTemp.Width = MS.CellWidth
            txtTemp.Top = MS.Top + MS.CellTop
            txtTemp.Left = MS.Left + MS.CellLeft
            txtTemp.Visible = True
            txtTemp = Chr(KeyAscii)
            txtTemp.SelStart = Len(Trim(txtTemp))
            txtTemp.SetFocus
        'Else
        '    MsgBox "Cannot Modify.Sale Transaction Exists", vbInformation + vbOKOnly, "DAS Version 4.0"
        '    MS.SetFocus
        'End If
    End If
End If
End Sub

Private Sub MSA_GotFocus()
If Ynlist.Visible = True Then Ynlist.Visible = False
End Sub

Private Sub MSA_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "y" Or Chr(KeyAscii) = "Y" Or Chr(KeyAscii) = "n" Or Chr(KeyAscii) = "N" Then
    If MSA.Col = 3 Then
        Ynlist.Visible = True
        Ynlist.Left = MSA.Left + MSA.CellLeft
        Ynlist.Top = MSA.Top + MSA.CellTop
        Ynlist.Height = MSA.CellHeight
        Ynlist.Width = MSA.CellWidth
        Ynlist.ListIndex = 0
        Ynlist.SetFocus
    End If
End If
End Sub

Private Sub MSA_LostFocus()
If SSTab1.Tab <> 2 Then
        MSA.Rows = 1
        txtMobile = ""
        txtInvoice = ""
        txtRetailer = ""
        txtSalesman = ""
End If
End Sub

Private Sub MSC_GotFocus()
If CYnlist.Visible = True Then CYnlist.Visible = False
End Sub

Private Sub msc_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "y" Or Chr(KeyAscii) = "Y" Or Chr(KeyAscii) = "n" Or Chr(KeyAscii) = "N" Then
    If MSC.Col = 3 Then
        CYnlist.Visible = True
        CYnlist.Left = MSC.Left + MSC.CellLeft
        CYnlist.Top = MSC.Top + MSC.CellTop
        CYnlist.Height = MSC.CellHeight
        CYnlist.Width = MSC.CellWidth
        CYnlist.ListIndex = 0
        CYnlist.SetFocus
    End If
End If
End Sub

Private Sub MSC_LostFocus()
If SSTab1.Tab <> 5 Then
MSC.Rows = 1
        txtcMobile = ""
        txtcInvoice = ""
        txtcRetailer = ""
        txtcSalesman = ""
        'cafdate.SetFocus
        'optDate.Value = True
        'chkCAll.SetFocus
        'chkCAll.Value = 1
        End If
End Sub

Private Sub MSPC_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
    If MSPC.Row > 0 Then
        If MSPC.Col = 3 Then
            Text2.Visible = True
            Text2.Left = MSPC.Left + MSPC.CellLeft
            Text2.Top = MSPC.Top + MSPC.CellTop
            Text2.Height = MSPC.CellHeight
            Text2.Width = MSPC.CellWidth
            Text2 = Chr(KeyAscii)
            Text2.SelStart = Len(Trim(Text2))
            Text2.SetFocus
        End If
    End If
End If
End Sub

Private Sub MSR_GotFocus()
If RYnlist.Visible = True Then RYnlist.Visible = False
End Sub

Private Sub MSR_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "y" Or Chr(KeyAscii) = "Y" Or Chr(KeyAscii) = "n" Or Chr(KeyAscii) = "N" Then
    If MSR.Col = 3 Then
        RYnlist.Visible = True
        RYnlist.Left = MSR.Left + MSR.CellLeft
        RYnlist.Top = MSR.Top + MSR.CellTop
        RYnlist.Height = MSR.CellHeight
        RYnlist.Width = MSR.CellWidth
        RYnlist.ListIndex = 0
        RYnlist.SetFocus
    End If
End If
End Sub

Private Sub optDate_Click()
txtcMobile = ""
txtcInvoice = ""
txtcRetailer = ""
txtcSalesman = ""
cafdate.SetFocus
optDate.Value = True
End Sub

Private Sub optDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cafdate.SetFocus
    txtcSalesman = ""
    txtcInvoice = ""
    txtcRetailer = ""
    txtcMobile = ""
    optDate.Value = True
End If
End Sub

Private Sub MSR_LostFocus()
        If SSTab1.Tab <> 6 Then
        MSR.Rows = 1
        txtrMobile = ""
        txtrInvoice = ""
        txtrRetailer = ""
        txtrSalesman = ""
        txtrMobile = ""
End If
End Sub

Private Sub optInvoice_Click()
If txtcInvoice.Enabled = False Then txtcInvoice.Enabled = True
txtcMobile = ""
txtcRetailer = ""
txtcSalesman = ""
txtcInvoice.SetFocus
End Sub

Private Sub optInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cafdate = Date$
    txtcSalesman = ""
    txtcInvoice.SetFocus
    txtcRetailer = ""
    txtcMobile = ""
End If
End Sub


Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
     If txtMobile.Enabled = False Then txtMobile.Enabled = True
        txtMobile.SetFocus
        txtInvoice = ""
        txtRetailer = ""
        txtSalesman = ""
    ElseIf Index = 1 Then
    If txtInvoice.Enabled = False Then txtInvoice.Enabled = True
        txtInvoice.SetFocus
        txtMobile = ""
        txtRetailer = ""
        txtSalesman = ""
    ElseIf Index = 2 Then
        actDate.SetFocus
        txtMobile = ""
        txtInvoice = ""
        txtRetailer = ""
        txtSalesman = ""
        'Option1(2).Value = True
        If chkAAll.Enabled = False Then chkAAll.Enabled = True
        chkAAll.SetFocus
        chkAAll.Value = 1
    ElseIf Index = 3 Then
        txtInvoice = ""
        txtMobile = ""
        txtSalesman = ""
        If txtRetailer.Enabled = False Then txtRetailer.Enabled = True
        txtRetailer.SetFocus
    ElseIf Index = 4 Then
        txtInvoice = ""
        txtMobile = ""
        txtRetailer = ""
        If txtSalesman.Enabled = False Then txtSalesman.Enabled = True
        txtSalesman.SetFocus
    End If
End Sub

Private Sub Option1_GotFocus(Index As Integer)
chkAAll.Value = 0
End Sub

Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    If Index = 0 Then
        txtMobile.SetFocus
        txtInvoice = ""
        txtRetailer = ""
        txtSalesman = ""
    ElseIf Index = 1 Then
        txtInvoice.SetFocus
        txtMobile = ""
        txtRetailer = ""
        txtSalesman = ""
    ElseIf Index = 2 Then
        actDate.SetFocus
        txtMobile = ""
        txtInvoice = ""
        txtRetailer = ""
        txtSalesman = ""
        'Option1(2).Value = True
        chkAAll.SetFocus
        chkAAll.Value = 1
    ElseIf Index = 3 Then
        txtInvoice = ""
        txtMobile = ""
        txtSalesman = ""
        txtRetailer.SetFocus
    ElseIf Index = 4 Then
        txtInvoice = ""
        txtMobile = ""
        txtRetailer = ""
        txtSalesman.SetFocus
    End If
End If
End Sub

Private Sub Option2_Click()
SRT = "SAL"
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtModRet.Enabled = True
    txtModRet.SetFocus
End If
End Sub
Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtModRet.Enabled = True
    txtModRet.SetFocus
End If
End Sub

Private Sub Option3_Click()
SRT = "OAP"
End Sub

Private Sub optMobile_Click()
If txtcMobile.Enabled = False Then txtcMobile.Enabled = True
txtcMobile.SetFocus
txtcInvoice = ""
txtcRetailer = ""
txtcSalesman = ""
End Sub

Private Sub optMobile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cafdate = Date$
    txtcSalesman = ""
    txtcInvoice = ""
    txtcRetailer = ""
    txtcMobile.SetFocus
End If
End Sub

Private Sub optrDate_Click()
txtrMobile = ""
txtrInvoice = ""
txtrRetailer = ""
txtrSalesman = ""
ChrDate.SetFocus
optrDate.Value = True
End Sub

Private Sub optrDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    ChrDate.SetFocus
End If
End Sub

Private Sub optRet_Click()
If txtcRetailer.Enabled = False Then txtcRetailer.Enabled = True
txtcMobile = ""
txtcInvoice = ""
txtcSalesman = ""
txtcRetailer.SetFocus
End Sub

Private Sub optRet_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cafdate = Date$
    txtcSalesman = ""
    txtcInvoice = ""
    txtcRetailer.SetFocus
    txtcMobile = ""
End If
End Sub

Private Sub optrInvoice_Click()
chkRAll.Value = 0
txtrMobile = ""
txtrRetailer = ""
txtrSalesman = ""
txtrInvoice.SetFocus
End Sub

Private Sub optrMobile_Click()
chkRAll.Value = 0
txtrMobile.SetFocus
txtrInvoice = ""
txtrRetailer = ""
txtrSalesman = ""
End Sub

Private Sub optrRet_Click()
chkRAll.Value = 0
txtrMobile = ""
txtrInvoice = ""
txtrSalesman = ""
txtrRetailer.SetFocus
End Sub

Private Sub optrSmn_Click()
chkRAll.Value = 0
txtrMobile = ""
txtrInvoice = ""
txtrRetailer = ""
txtrSalesman.SetFocus
End Sub


Private Sub optSmn_Click()
txtcMobile = ""
txtcInvoice = ""
txtcRetailer = ""
txtcSalesman = ""
End Sub

Private Sub optSmn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cafdate = Date$
    txtcSalesman.SetFocus
    txtcInvoice = ""
    txtcRetailer = ""
    txtcMobile = ""
End If
End Sub


Private Sub RYnlist_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSR.Text = RYnlist.Text
    RYnlist.Visible = False
    If MSR.Row <> MSR.Rows - 1 Then MSR.Row = MSR.Row + 1
    MSR.SetFocus
End If
End Sub

Private Sub salimpcmd_Click()
If txtSalInv.Text <> "" Then
On Error GoTo errh
Dim pno As Integer
cd1.ShowOpen
Open cd1.FileName For Input As #1
Line Input #1, RLINE
adate = Trim(RLINE)
pno = 0
Do Until EOF(1)
    Line Input #1, RLINE
    SMBL = Mid(Trim(RLINE), 1, 10)
    pno = pno + 1
Loop
Close #1

If pno = lv.ListItems.count Then
Open cd1.FileName For Input As #1
Line Input #1, RLINE
     
     Do Until EOF(1)
        Line Input #1, RLINE
        txtSMbl.Text = Trim(Mid(Trim(RLINE), 1, 10))
        txtSMbl_KeyDown 13, 0
     Loop
    Close #1
   
   ' MsgBox "Invalid Mobile Numbers SpreadSheet File Selected!", vbCritical + vbOKOnly, "DAS Version 4.0"
    Exit Sub
End If

Exit Sub
errh:
MsgBox "Select a Valid File!", vbInformation, "DAS Version 4.0"
Else
    MsgBox "Invoice not Selected!", vbInformation, "DAS Version 4.0"
End If
End Sub
Private Sub sal_imp_Click()

Dim appXL As Object
Dim appX As Object
Dim TYP As String
Set appXL = CreateObject("Excel.Application")
CommonDialog1.FileName = ""
CommonDialog1.filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
CommonDialog1.DialogTitle = "Select Mobile Numbers File"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen
CommonDialog1.CancelError = False
If Not Trim(CommonDialog1.FileName) = "" Then
    If LCase(Right(Trim(CommonDialog1.FileName), 3)) = "xls" Then ' Edited
         appXL.Workbooks.Open (CommonDialog1.FileName)  'GetObject(Trim(CommonDialog1.FileName))
    Else
        MsgBox "Selected File is not a Valid SpreadSheet File"
        Exit Sub
    End If
    TYP = appXL.Worksheets(1).Cells(1, 1)
    If TYP = "Purchase" Then
        Dim RSTEMP As New ADODB.Recordset
        Dim rstemp1 As New ADODB.Recordset
        Dim rstemp2 As New ADODB.Recordset
        Dim rstemp3 As New ADODB.Recordset
        Dim pno1 As Integer
        On Error GoTo errh
        Dim mn() As String
        Dim mn1() As String
        Dim chk_flg As Boolean
        chk_flg = True
        Dim pno As Integer
        Dim row_num As Integer
        If RSTEMP.State = 1 Then RSTEMP.Close
        RSTEMP.Open "select chr_one,chr_two from rp2", con, adOpenKeyset, adLockPessimistic
        row_num = 3
        tot_num = 0
        While Len(Trim(appXL.Worksheets(1).Cells(row_num, 2))) > 0
            If Len(Trim(Val(appXL.Worksheets(1).Cells(row_num, 2)))) = 10 And Trim(appXL.Worksheets(1).Cells(row_num, 5)) <> "" And Trim(appXL.Worksheets(1).Cells(row_num, 6)) <> "" Then
                     If rstemp2.State = 1 Then rstemp2.Close
                rstemp2.Open "select MBL_IDY,SAL_IDY,SAL_DAT,SAL_TYP,CST_IDY,SMN_IDY,PRD_IDY from avt WHERE mbl_idy='" & Trim(appXL.Worksheets(1).Cells(row_num, 2)) & "'AND SAL_TYP IS NULL", con, adOpenDynamic, adLockPessimistic
                     If rstemp2.EOF = False Then
                        chk_flg = CHK_SAL_IMP(Trim(appXL.Worksheets(1).Cells(row_num, 5)), Trim(appXL.Worksheets(1).Cells(row_num, 4)))
                        If chk_flg = True Then
                        Dim RSTEMP4 As New ADODB.Recordset
                        If RSTEMP4.State = 1 Then RSTEMP4.Close
                        RSTEMP4.Open "select PRD_IDY from avt WHERE mbl_idy='" & Trim(appXL.Worksheets(1).Cells(row_num, 2)) & "'AND SAL_TYP IS NULL", con, adOpenDynamic, adLockPessimistic
                        Dim PRDRS As New ADODB.Recordset
                        If PRDRS.State = 1 Then PRDRS.Close
                        PRDRS.Open "SELECT PRD_NME, ALT_IDY FROM Prd where prd_IDY like '" & RSTEMP4("prd_idy") & "'", con, adOpenDynamic, adLockPessimistic
                        If Trim(PRDRS("PRD_NME")) = Trim(appXL.Worksheets(1).Cells(row_num, 3)) And Trim(PRDRS("ALT_IDY")) = Trim(appXL.Worksheets(1).Cells(row_num, 4)) Then
                        If row_num < 4 Then
                            rstemp2("sal_idy") = Trim(appXL.Worksheets(1).Cells(row_num, 5))
                            rstemp2("sal_dat") = Trim(appXL.Worksheets(1).Cells(row_num, 6))
                            rstemp2("sal_typ") = "SAL"
                            If rstemp3.State = 1 Then rstemp3.Close
                            rstemp3.Open "select CST_IDY,SMN_IDY from INL WHERE INV_idy='" & Trim(appXL.Worksheets(1).Cells(row_num, 5)) & "' AND TXN_TYP='SAL'", con, adOpenDynamic, adLockPessimistic
                            rstemp2("cst_idy") = Trim(rstemp3(0))
                            rstemp2("smn_idy") = Trim(rstemp3(1))
                            rstemp2.Update
                        ElseIf row_num > 3 Then
                            Dim INVDATE As Date
                            
                            If rstemp3.State = 1 Then rstemp3.Close
                            rstemp3.Open "select CST_IDY,SMN_IDY,INV_DAT from INL WHERE INV_idy='" & Trim(appXL.Worksheets(1).Cells(row_num, 5)) & "' AND TXN_TYP='SAL'", con, adOpenDynamic, adLockPessimistic
                            INVDATE = Format(rstemp3("INV_DAT"), "MM/DD/YYYY")
                            If INVDATE = Trim(appXL.Worksheets(1).Cells(row_num, 6)) Then
                            rstemp2("sal_idy") = Trim(appXL.Worksheets(1).Cells(row_num, 5))
                            rstemp2("sal_dat") = Trim(appXL.Worksheets(1).Cells(row_num, 6))
                            rstemp2("sal_typ") = "SAL"
                            rstemp2("cst_idy") = Trim(rstemp3(0))
                            rstemp2("smn_idy") = Trim(rstemp3(1))
                            rstemp2.Update
                            Else
                            
                            End If
                        End If
                        End If
                         tot_num = Val(tot_num) + 1
                    End If
                    End If
                End If
            row_num = Val(row_num) + 1
        Wend
 
        MsgBox "Sale Invoice Data is updated!", vbInformation, "DAS Version 4.0"
        lv.ListItems.CLEAR
       CLEAR
        For i = 1 To appXL.Workbooks.count
        appXL.Workbooks(i).Close
        Next
        Set appXL = Nothing
        Exit Sub

End If
    MsgBox "Invalid File!", vbCritical, "DAS Version 4.0"
    appXL.Workbooks.Close
    Set appXL = Nothing
    Exit Sub
Else
    MsgBox "Mobile Numbers File Selection Cancelled!"
    Exit Sub
End If
errh:
    MsgBox err.Description
End Sub





'End Sub
'Dim appXL As Object
'Dim appX As Object
'Dim TYP As String
'Set appXL = CreateObject("Excel.Application")
'CommonDialog1.FileName = ""
'CommonDialog1.filter = "All Files (*.*)|*.*|Excel Files (*.xls)|*.xls"
'CommonDialog1.DialogTitle = "Select Mobile Numbers File"
'CommonDialog1.FilterIndex = 2
'CommonDialog1.ShowOpen
'CommonDialog1.CancelError = False
'If Not Trim(CommonDialog1.FileName) = "" Then
'    If LCase(Right(Trim(CommonDialog1.FileName), 3)) = "xls" Then ' Edited
'         appXL.Workbooks.Open (CommonDialog1.FileName)  'GetObject(Trim(CommonDialog1.FileName))
'    Else
'        MsgBox "Selected File is not a Valid SpreadSheet File"
'        Exit Sub
'    End If
'    TYP = appXL.Worksheets(1).Cells(1, 1)
'    If TYP = "Purchase" Then
'        Dim RSTEMP As New ADODB.Recordset
'        Dim rstemp1 As New ADODB.Recordset
'        Dim rstemp2 As New ADODB.Recordset
'        Dim rstemp3 As New ADODB.Recordset
'        Dim saledate As Date
'        Dim pno1 As Integer
'        On Error GoTo errh
'        Dim mn() As String
'        Dim mn1() As String
'        Dim chk_flg As Boolean
'        chk_flg = True
'        Dim pno As Integer
'        Dim row_num As Integer
'        If RSTEMP.State = 1 Then RSTEMP.Close
'        RSTEMP.Open "select chr_one,chr_two from rp2", CON, adOpenKeyset, adlockpessimistic
'        row_num = 3
'        tot_num = 0
'        While Len(Trim(appXL.Worksheets(1).Cells(row_num, 2))) > 0
'            If Len(Trim(Val(appXL.Worksheets(1).Cells(row_num, 2)))) = 10 And Trim(appXL.Worksheets(1).Cells(row_num, 5)) <> "" And Trim(appXL.Worksheets(1).Cells(row_num, 6)) <> "" Then
'                     If rstemp2.State = 1 Then rstemp2.Close
'                rstemp2.Open "select MBL_IDY,SAL_IDY,SAL_DAT,SAL_TYP,CST_IDY,SMN_IDY,PRD_IDY from avt WHERE mbl_idy='" & Trim(appXL.Worksheets(1).Cells(row_num, 2)) & "'AND SAL_TYP IS NULL", CON, adOpenDynamic, adlockpessimistic
'                     If rstemp2.EOF = False Then
'                        chk_flg = CHK_SAL_IMP(Trim(appXL.Worksheets(1).Cells(row_num, 5)), Trim(appXL.Worksheets(1).Cells(row_num, 4)))
'                        If chk_flg = True Then
'                        Dim RSTEMP4 As New ADODB.Recordset
'                        If RSTEMP4.State = 1 Then RSTEMP4.Close
'                        RSTEMP4.Open "select PRD_IDY from avt WHERE mbl_idy='" & Trim(appXL.Worksheets(1).Cells(row_num, 2)) & "'AND SAL_TYP IS NULL", CON, adOpenDynamic, adlockpessimistic
'                        Dim PRDRS As New ADODB.Recordset
'                        If PRDRS.State = 1 Then PRDRS.Close
'                        PRDRS.Open "SELECT PRD_NME, ALT_IDY FROM Prd where prd_IDY like '" & RSTEMP4("prd_idy") & "'", CON, adOpenDynamic, adlockpessimistic
'                        If Trim(PRDRS("PRD_NME")) = Trim(appXL.Worksheets(1).Cells(row_num, 3)) And Trim(PRDRS("ALT_IDY")) = Trim(appXL.Worksheets(1).Cells(row_num, 4)) Then
'                        If row_num < 4 Then
'                            rstemp2("sal_idy") = Trim(appXL.Worksheets(1).Cells(row_num, 5))
'                            saledate = Trim(appXL.Worksheets(1).Cells(row_num, 6))
'                            saledate = Format(saledate, "mm / dd / yyyy") 'changed by pkb for dateformat
'                            saledate = Format(saledate, "dd / mm / yyyy")
'                            rstemp2("sal_dat") = Trim(appXL.Worksheets(1).Cells(row_num, 6))
'                             rstemp2("sal_dat") = saledate
'                            rstemp2("sal_typ") = "SAL"
'                            If rstemp3.State = 1 Then rstemp3.Close
'                            rstemp3.Open "select CST_IDY,SMN_IDY from INL WHERE INV_idy='" & Trim(appXL.Worksheets(1).Cells(row_num, 5)) & "' AND TXN_TYP='SAL'", CON, adOpenDynamic, adlockpessimistic
'                            rstemp2("cst_idy") = Trim(rstemp3(0))
'                            rstemp2("smn_idy") = Trim(rstemp3(1))
'                            rstemp2.Update
'                        ElseIf row_num > 3 Then
'                            Dim INVDATE As Date
'
'                            If rstemp3.State = 1 Then rstemp3.Close
'                            rstemp3.Open "select CST_IDY,SMN_IDY,INV_DAT from INL WHERE INV_idy='" & Trim(appXL.Worksheets(1).Cells(row_num, 5)) & "' AND TXN_TYP='SAL'", CON, adOpenDynamic, adlockpessimistic
'                        '    INVDATE = Format(rstemp3("INV_DAT"), "MM/DD/YYYY")
'                            saledate = Trim(appXL.Worksheets(1).Cells(row_num, 6))
'                            saledate = Format(saledate, "mm / dd / yyyy") 'changed by pkb for dateformat
'                            saledate = Format(saledate, "dd / mm / yyyy")
'                           ' If INVDATE = Trim(appXL.Worksheets(1).Cells(row_num, 6)) Then
'                           If rstemp3("INV_DAT") = saledate Then
'                            rstemp2("sal_idy") = Trim(appXL.Worksheets(1).Cells(row_num, 5))
'                           ' rstemp2("sal_dat") = Trim(appXL.Worksheets(1).Cells(row_num, 6))
'                           rstemp2("SAL_DAT ") = saledate
'                            rstemp2("sal_typ") = "SAL"
'                            rstemp2("cst_idy") = Trim(rstemp3(0))
'                            rstemp2("smn_idy") = Trim(rstemp3(1))
'                            rstemp2.Update
'                            Else
'
'                            End If
'                        End If
'                        End If
'                         tot_num = Val(tot_num) + 1
'                    End If
'                    End If
'                End If
'            row_num = Val(row_num) + 1
'        Wend
'
'        MsgBox "Sale Invoice Data is updated!", vbInformation, "DAS Version 4.0"
'        lv.ListItems.CLEAR
'       CLEAR
'        For I = 1 To appXL.Workbooks.COUNT
'        appXL.Workbooks(I).Close
'        Next
'        Set appXL = Nothing
'        Exit Sub
'
'End If
'    MsgBox "Invalid File!", vbCritical, "DAS Version 4.0"
'    appXL.Workbooks.Close
'    Set appXL = Nothing
'    Exit Sub
'Else
'    MsgBox "Mobile Numbers File Selection Cancelled!"
'    Exit Sub
'End If
'errh:
'    MsgBox err.Description
'End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.Tab = 0 Then
    If MS.Rows > 1 Then
        MS.Row = 1
        MS.Col = 3
        MS.SetFocus
    Else
        txtVch.SetFocus
    End If
ElseIf SSTab1.Tab = 1 Then

    If lv.ListItems.count >= 1 Then
        lv.SetFocus
    Else
        txtSalInv.SetFocus
    End If
ElseIf SSTab1.Tab = 2 Then
    MSA.Cols = 7
    MSA.TextMatrix(0, 0) = "SNO"
    MSA.ColWidth(0) = 700
    MSA.TextMatrix(0, 1) = "Mobile ID"
    MSA.ColWidth(1) = 1300
    MSA.ColAlignment(1) = 1
    MSA.TextMatrix(0, 2) = "Invoice ID"
    MSA.ColWidth(2) = 1300
    MSA.TextMatrix(0, 3) = "Status"
    MSA.ColWidth(3) = 800
    MSA.TextMatrix(0, 4) = "SKU Name"
    MSA.ColWidth(4) = 2200
    MSA.TextMatrix(0, 5) = "Party"
    MSA.ColWidth(5) = 1740
    MSA.TextMatrix(0, 6) = "Invoice Date"
    MSA.ColWidth(6) = 1200
    If MSA.Rows > 1 Then
        If Len(Trim(MSA.TextMatrix(1, 1))) > 0 Then
            MSA.Col = 3
            MSA.SetFocus
        Else
            chkAAll.Enabled = True
            chkAAll.SetFocus
            chkAAll.Value = 1
            Option1.Item(1).Value = False
            Option1.Item(3).Value = False
            Option1.Item(4).Value = False
            Option1.Item(0).Value = False
            Option1.Item(1).Enabled = True
            Option1.Item(3).Value = False
            Option1.Item(4).Value = False
            Option1.Item(0).Value = False
        End If
    Else
        If chkAAll.Enabled = False Then chkAAll.Enabled = True
        chkAAll.SetFocus
        chkAAll.Value = 1
        Option1.Item(1).Value = False
        Option1.Item(3).Value = False
        Option1.Item(4).Value = False
        Option1.Item(0).Value = False
    End If
    For s = 0 To 5
    If Ptno(s).Visible = True Then Ptno(s).Visible = False
    Next
ElseIf SSTab1.Tab = 4 Then
    cmbType.ListIndex = 0
    LOADMSPC
    If LVF.ListItems.count > 0 Then
        If LVF.Visible = True Then LVF.SetFocus
    Else
        cmbType.SetFocus
    End If
ElseIf SSTab1.Tab = 5 Then
'    chkCAll.Enabled = True
'    chkCAll.Value = 1
'   chkCAll.SetFocus
    MSC.Cols = 7
    MSC.TextMatrix(0, 0) = "SNO"
    MSC.ColWidth(0) = 1300
    MSC.TextMatrix(0, 1) = "Mobile ID"
    MSC.ColWidth(1) = 1300
    MSC.ColAlignment(1) = 1
    MSC.TextMatrix(0, 2) = "Invoice ID"
    MSC.ColWidth(2) = 1300
    MSC.TextMatrix(0, 3) = "Status"
    MSC.ColWidth(3) = 800
    MSC.TextMatrix(0, 4) = "SKU Name"
    MSC.ColWidth(4) = 2200
    MSC.TextMatrix(0, 5) = "Party"
    MSC.ColWidth(5) = 1740
    MSC.TextMatrix(0, 6) = "Invoice Date"
    MSC.ColWidth(6) = 1200
    If MSC.Rows > 1 Then
        If Len(Trim(MSC.TextMatrix(1, 0))) > 0 Then
            MSC.Col = 3
            MSC.SetFocus
        Else
            chkCAll.Enabled = True
            chkCAll.SetFocus
            chkCAll.Value = 1
            optSmn.Value = False
            optRet.Value = False
            optInvoice.Value = False
            optMobile.Value = False
        End If
    Else
            chkCAll.Enabled = True
            chkCAll.SetFocus
            chkCAll.Value = 1
            optSmn.Value = False
            optRet.Value = False
            optInvoice.Value = False
            optMobile.Value = False
    End If
'            'optDate.Value = True
'            'optDate.SetFocus
''            chkCAll.SetFocus
''            chkCAll.Value = 1
''            Option1.Item(1).Value = False
''            Option1.Item(3).Value = False
''            Option1.Item(4).Value = False
''            Option1.Item(0).Value = False
'        If optSmn.Value = False Then
'        If optInvoice.Value = False Then
'        If optRet.Value = False Then
'        If optMobile.Value = False Then
'        chkCAll.Enabled = True
'        chkCAll.SetFocus
'        chkCAll.Value = 1
'        End If
'        End If
'        End If
'        End If
'        End If
'    Else
'        'optDate.Value = True
'        'optDate.SetFocus
'        If optSmn.Value = False Then
'        If optInvoice.Value = False Then
'        If optRet.Value = False Then
'        If optMobile.Value = False Then
'        chkCAll.SetFocus
'        chkCAll.Value = 1
'        End If
'        End If
'        End If
'        End If
'
'    End If
ElseIf SSTab1.Tab = 6 Then
'    chkRAll.Enabled = True
'    chkRAll.Value = 1
'    chkRAll.SetFocus
    MSR.Cols = 7
    MSR.TextMatrix(0, 0) = "SNO"
    MSR.ColWidth(0) = 700
    MSR.TextMatrix(0, 1) = "Mobile ID"
    MSR.ColWidth(1) = 1300
    MSR.ColAlignment(1) = 1
    MSR.TextMatrix(0, 2) = "Invoice ID"
    MSR.ColWidth(2) = 1300
    MSR.TextMatrix(0, 3) = "Status"
    MSR.ColWidth(3) = 800
    MSR.TextMatrix(0, 4) = "SKU Name"
    MSR.ColWidth(4) = 2200
    MSR.TextMatrix(0, 5) = "Party"
    MSR.ColWidth(5) = 1740
    MSR.TextMatrix(0, 6) = "Invoice Date"
    MSR.ColWidth(6) = 1200
    If MSR.Rows > 1 Then
        If Len(Trim(MSR.TextMatrix(1, 0))) > 0 Then
            MSR.Col = 3
            MSR.SetFocus
        Else
            chkRAll.Enabled = True
            chkRAll.SetFocus
            chkRAll.Value = 1
            optrSmn.Value = False
            optrRet.Value = False
            optrInvoice.Value = False
            optrMobile.Value = False
        End If
    Else
            chkRAll.Enabled = True
            chkRAll.SetFocus
            chkRAll.Value = 1
            optrSmn.Value = False
            optrRet.Value = False
            optrInvoice.Value = False
            optrMobile.Value = False
    End If
            'optrDate.Value = True
            'optrDate.SetFocus
'        If optrSmn.Value = False Then
'        If optrRet.Value = False Then
'        If optrInvoice.Value = False Then
'        If optrMobile.Value = False Then
'            chkRAll.SetFocus
'            chkRAll.Value = 1
'        End If
'        End If
'        End If
'        End If
 '   End If
   ' Else
        'optrDate.Value = True
        'optrDate.SetFocus
'        If optrSmn.Value = False Then
'        If optrRet.Value = False Then
'        If optrInvoice.Value = False Then
'        If optrMobile.Value = False Then
'            chkRAll.SetFocus
'            chkRAll.Value = 1
'        End If
'        End If
'        End If
'        End If
   ' End If
ElseIf SSTab1.Tab = 7 Then
    chkModDate.Enabled = True
    chkModDate.SetFocus
'    lblno(0).Visible = False
'    lblno(3).Visible = False
'    lblno(2).Visible = False
'    lblno(1).Visible = False
    'chkModDate.Value = 1
End If
End Sub

Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(Val(Text2))) < 10 Then
        MsgBox "10 Digit Mobile Number Required!", vbInformation + vbOKOnly, "DAS Version 4.0"
        Text2.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If rschk.State = 1 Then rschk.Close
    rschk.Open "SELECT mbl_idy FROM AVT WHERE mbl_idy = '" & Trim(Text2) & "'", con, adOpenKeyset, adLockPessimistic
  
    'and MBL_IDY <> '" & Trim(MSPC.Text) & "'", CON, adOpenKeyset, adlockpessimistic
    If rschk.EOF = True Then
        If MsgBox("Update Mobile Number ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
            If rsEdt.State = 1 Then rsEdt.Close
            rsEdt.Open "select mbl_idy from avt where mbl_idy='" & Trim(MSPC.Text) & "'", con, adOpenKeyset, adLockPessimistic
            rsEdt("mbl_idy") = Trim(Text2)
            rsEdt.Update
            MSPC.Text = Trim(Text2)
            Text2.Visible = False
            If MSPC.Row <> MSPC.Rows - 1 Then
                MSPC.Row = MSPC.Row + 1
                MSPC.Col = 3
            End If
'            MSPC.Row = 1
'            MSPC.Col = 3
            MSPC.SetFocus
        End If
    Else
        MsgBox "Mobile Number Already Exists!", vbInformation + vbOKOnly, "DAS Version 4.0"
        Text2.Visible = False
        MSPC.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 42 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim chk As Integer, CHK1 As Integer
CHK1 = 0
If KeyCode = 13 Then
    Dim ALL_CHK As Boolean
    Dim temp As String
    Dim temp1 As String
    Dim temp2 As String
    Dim temp3 As String
    Dim temp4 As String
    Dim temp5 As String
    Dim temp6 As String
    Dim temp_chk As Integer
    
    ALL_CHK = False
    For i = 1 To LVR.ListItems.count
    If Trim(LVR.ListItems(i).Text) = Trim(txtSMbl.Text) Then
    ALL_CHK = True
    For i1 = 1 To LVR.ListItems.count
    If LVR.ListItems(i1).Checked = False Then
        temp_chk = i1
        Exit For
    End If
    Next
    
    
    temp = LVR.ListItems(i)
    temp1 = LVR.ListItems(i).SubItems(1)
    temp2 = LVR.ListItems(i).SubItems(2)
    temp3 = LVR.ListItems(i).SubItems(3)
    temp4 = LVR.ListItems(i).SubItems(4)
    temp5 = LVR.ListItems(i).SubItems(5)
    temp6 = LVR.ListItems(i).SubItems(6)
    
    For K = i To temp_chk + 1 Step -1
    LVR.ListItems(K).Text = LVR.ListItems(K - 1).Text
    LVR.ListItems(K).SubItems(1) = LVR.ListItems(K - 1).SubItems(1)
    LVR.ListItems(K).SubItems(2) = LVR.ListItems(K - 1).SubItems(2)
    LVR.ListItems(K).SubItems(3) = LVR.ListItems(K - 1).SubItems(3)
    LVR.ListItems(K).SubItems(4) = LVR.ListItems(K - 1).SubItems(4)
    LVR.ListItems(K).SubItems(5) = LVR.ListItems(K - 1).SubItems(5)
    LVR.ListItems(K).SubItems(6) = LVR.ListItems(K - 1).SubItems(6)
    Next
    
        LVR.ListItems(temp_chk).Text = temp
        LVR.ListItems(temp_chk).SubItems(1) = temp1
        LVR.ListItems(temp_chk).SubItems(2) = temp2
        LVR.ListItems(temp_chk).SubItems(3) = temp3
        LVR.ListItems(temp_chk).SubItems(4) = temp4
        LVR.ListItems(temp_chk).SubItems(5) = temp5
        LVR.ListItems(temp_chk).SubItems(6) = temp6
        LVR.ListItems(temp_chk).Checked = True
        LVR.ListItems(temp_chk).Selected = True
        CHK1 = 1
        Exit For
    End If
    Next
    
    If ALL_CHK = False Then
        MsgBox "Mobile Number Invalid! No Such Number", vbCritical, "DAS Version 4.0"
        txtSMbl.SetFocus
    Else
        LVR.SetFocus
    End If
End If



End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.KEY = "MBL_SUM" Then
    M_FRM_SAQ.Tag = "MBL_SUM"
    M_FRM_SAQ.Show 1
ElseIf ButtonMenu.KEY = "INW_REG" Then
    M_FRM_SAQ.Tag = "INW_REG"
    M_FRM_SAQ.Show 1
ElseIf ButtonMenu.KEY = "OUT_REG" Then
    M_FRM_SAQ.Tag = "OUT_REG"
    M_FRM_SAQ.Show 1
ElseIf ButtonMenu.KEY = "AVT_RPT" Then
    M_FRM_SAQ.Tag = "AVT_RPT"
    M_FRM_SAQ.Show 1
ElseIf ButtonMenu.KEY = "NAV_REP" Then
    M_FRM_SAQ.Tag = "NAV_REP"
    M_FRM_SAQ.Show 1
ElseIf ButtonMenu.KEY = "CAF_REP" Then
    M_FRM_SAQ.Tag = "CAF_REP"
    M_FRM_SAQ.Show 1
ElseIf ButtonMenu.KEY = "CPN_REP" Then
    M_FRM_SAQ.Tag = "CPN_REP"
    M_FRM_SAQ.Show 1
ElseIf ButtonMenu.KEY = "CHN_REP" Then
    M_FRM_SAQ.Tag = "CHN_REP"
    M_FRM_SAQ.Show 1
ElseIf ButtonMenu.KEY = "MBL_STK" Then
    'M_FRM_SAQ.Tag = "MBL_STK"
    'M_FRM_SAQ.Show 1
    OpenReport CURDIR & "\TCM.RP1", "Mobile Stock on Hand", 2, "(AVT->SAL_IDY=''.AND.AVT->TXN_TYP<>'MRS')", 6
ElseIf ButtonMenu.KEY = "MBL_STS" Then
    M_REP_MBL.Show 1
End If
End Sub

Private Sub txtcInvoice_GotFocus()
'code added by pkb
optInvoice.Value = True
txtcSalesman = ""
txtcRetailer = ""
txtcMobile = ""
If txtcInvoice = "" Then
MSC.CLEAR
End If
End Sub

Private Sub txtcInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    txtcInvoice = ""
    optMobile.Value = True
    txtcMobile.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    txtcInvoice = ""
    optRet.Value = True
    txtcRetailer.SetFocus
    
    ElseIf KeyCode = 37 And Shift = 1 Then
    txtcInvoice = ""
    txtcSalesman.SetFocus
    optSmn.Value = True
    
End If
If KeyCode = 13 Then
    Dim ls As ListItem
    Label49 = "Select Invoice for Mobile(s) Activation for CAF"
    SSTab1.TabEnabled(3) = True
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "select sal_idy,avt_dat,cst_idy,prd_idy,txn_typ from avt where SAL_IDY IS NOT NULL AND sal_typ like 'SAL' and avt_sts='Yes' and (caf_sts is null or caf_sts='' or caf_sts<>'Yes') group BY sal_idy", con, adOpenKeyset, adLockPessimistic
    If IsNull(RSINL(0)) = True Then
        MsgBox "Cannot Activate from Sales Invoice!", vbInformation + vbOKOnly, "DAS Version 4.0"
        txtcInvoice.SetFocus
    Else
    If RSINL.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
        LVP.ColumnHeaders.Add , , "Invoice ID"
        LVP.ColumnHeaders.Add , , "SKU Name", 2500
        LVP.ColumnHeaders.Add , , "Party", 2500
        LVP.ColumnHeaders.Add , , "Act. Date", 1200
        LVP.ListItems.CLEAR
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_idy,prd_nme from prd", con, adOpenKeyset, adLockPessimistic
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
        While Not RSINL.EOF
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_idy,prd_nme from prd where prd_idy='" & RSINL(3) & "'", con, adOpenKeyset, adLockPessimistic
            If IsNull(RSINL(0)) = True Then
            Set ls = LVP.ListItems.Add(, , " ")
            Else
            Set ls = LVP.ListItems.Add(, , RSINL(0))
            End If
            If Not rsprd.EOF Then ls.ListSubItems.Add , , rsprd(1)
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act where act_idy='" & RSINL(2) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then ls.ListSubItems.Add , , rsact(1)
            ls.ListSubItems.Add , , RSINL(1)
            RSINL.MoveNext
        Wend
        SSTab1.TabEnabled(3) = True
        LVP.Visible = True
        LVP.SetFocus
        SSTab1.Tab = 3
    Else
        MsgBox "Invoice ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
        SSTab1.Tab = 5
        txtcInvoice.SetFocus
        SSTab1.TabEnabled(3) = False
    End If
    End If
End If
End Sub

Private Sub txtcMobile_GotFocus()
'code added by pkb
optMobile.Value = True
        txtcSalesman = ""
        txtcRetailer = ""
        txtcInvoice = ""
        If txtcMobile = "" Then
        MSC.CLEAR
        End If
End Sub

Private Sub txtcMobile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 And Shift = 1 Then
    txtcMobile = ""
    optInvoice.Value = True
    txtcInvoice.SetFocus
ElseIf KeyCode = 37 And Shift = 1 Then
    txtcMobile = ""
    If txtcRetailer.Enabled = False Then txtcRetailer.Enabled = True
    txtcRetailer.SetFocus
    optRet.Value = True
End If
If KeyCode = 13 Then
If Not Len(Trim(txtcMobile)) > 9 Then
    Dim ls As ListItem
    Label49 = "Select Mobile NO for CAF"
    SSTab1.TabEnabled(3) = True
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts,caf_sts from avt where sal_typ='SAL' AND avt_sts='Yes' and (caf_sts is null or caf_sts<>'Yes' or caf_sts='') and mbl_idy like '" & Trim(txtcMobile) & "%'", con, adOpenKeyset, adLockPessimistic
    If RSINL.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
        LVP.ColumnHeaders.Add , , "Mobile ID"
        LVP.ColumnHeaders.Add , , "Invoice ID"
        LVP.ColumnHeaders.Add , , "SKU Name"
        LVP.ColumnHeaders.Add , , "Party"
        LVP.ColumnHeaders.Add , , "Act. Date"
        LVP.ListItems.CLEAR
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
        While Not RSINL.EOF
            If rsprd.RecordCount > 0 Then rsprd.MoveFirst
            Set ls = LVP.ListItems.Add(, , RSINL(0))
            If IsNull(RSINL(1)) = True Then
            ls.ListSubItems.Add , , " "
            Else
            ls.ListSubItems.Add , , RSINL(1)
            End If
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsprd.EOF Then ls.ListSubItems.Add , , Trim(rsprd(0))
            If rsact.RecordCount > 0 Then rsact.MoveFirst
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act where act_idy='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then ls.ListSubItems.Add , , Trim(rsact(1)) & ""
            ls.ListSubItems.Add , , RSINL(2)
            RSINL.MoveNext
        Wend
        LVP.Visible = True
        LVP.SetFocus
        SSTab1.Tab = 3
    Else
        MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
        SSTab1.Tab = 5
        txtcMobile.SetFocus
    End If
Else
    If Len(Trim(txtcMobile)) > 9 Then
      MBLID = Trim(txtcMobile)
      '  MBLID = LVP.SelectedItem.Text
        If RSINL.State = 1 Then RSINL.Close
        Set RSINL = Nothing
        RSINL.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts,caf_sts from avt where sal_typ='SAL' AND avt_sts='Yes' and (caf_sts is null or caf_sts<>'Yes' or caf_sts='') and mbl_idy='" & Trim(MBLID) & "'", con, adOpenKeyset, adLockPessimistic
        If RSINL.RecordCount > 0 Then
            MSC.Rows = 1
            MSC.Cols = 7
            MSC.TextMatrix(0, 0) = "SNO"
            MSC.ColWidth(0) = 700
            MSC.TextMatrix(0, 1) = "Mobile ID"
            MSC.ColWidth(1) = 1300
            MSC.ColAlignment(1) = 1
            MSC.TextMatrix(0, 2) = "Invoice ID"
            MSC.ColWidth(2) = 1300
            MSC.TextMatrix(0, 3) = "Status"
            MSC.ColWidth(3) = 800
            MSC.TextMatrix(0, 4) = "SKU Name"
            MSC.ColWidth(4) = 2200
            MSC.TextMatrix(0, 5) = "Party"
            MSC.ColWidth(5) = 1740
            MSC.TextMatrix(0, 6) = "Act. Date"
            MSC.ColWidth(6) = 1200

            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
            MSC.Rows = RSINL.RecordCount + 1
            r = 1
            While Not RSINL.EOF
                MSC.TextMatrix(r, 0) = r
                MSC.TextMatrix(r, 1) = Trim(RSINL(0))
                If IsNull(RSINL(1)) = True Then
                MSC.TextMatrix(r, 2) = ""
                Else
                MSC.TextMatrix(r, 2) = Trim(RSINL(1))
                End If
                txtcMobile.Text = Trim(RSINL(0))
                MSC.TextMatrix(r, 3) = "No" 'Trim(RSINL(7)) & ""
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsprd.EOF Then MSC.TextMatrix(r, 4) = Trim(rsprd(0))
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsact.EOF Then MSC.TextMatrix(r, 5) = Trim(rsact(1))
                MSC.TextMatrix(r, 6) = Trim(RSINL(2))
                r = r + 1
                RSINL.MoveNext
            Wend
            SSTab1.Tab = 5
            SSTab1.TabEnabled(3) = False
            MSC.Col = 3
            MSC.SetFocus
        Else
            MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
            SSTab1.Tab = 5
            'cafdate.SetFocus
            'optDate.Value = True
            If chkCAll.Enabled = False Then chkCAll.Enabled = True
            chkCAll.SetFocus
            chkCAll.Value = 1
            SSTab1.TabEnabled(3) = False
        End If
        End If
End If

End If
End Sub

Private Sub txtcMobile_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub txtcRetailer_GotFocus()
'code added by pkb
    optRet.Value = True
        txtcInvoice = ""
        txtcSalesman = ""
        txtcMobile = ""
        If txtcRetailer = "" Then
        MSC.CLEAR
        End If
End Sub

Private Sub txtcRetailer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    txtcRetailer = ""
    optInvoice.Value = True
    txtcInvoice.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    txtcRetailer = ""
    optSmn.Value = True
    txtcSalesman.SetFocus
ElseIf KeyCode = 39 And Shift = 1 Then
    txtcRetailer = ""
    txtcMobile.SetFocus
    optMobile.Value = True
End If
If KeyCode = 13 Then
If Not Val(txtcRetailer) > 0 Then
    Dim ls As ListItem
    Label49 = "Select Party Account for CAF"
    SSTab1.TabEnabled(3) = True
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "select sal_idy,avt_dat,cst_idy,prd_idy,txn_typ from avt where sal_typ like 'SAL' and avt_sts='Yes' and (caf_sts is null or caf_sts='' or caf_sts<>'Yes') group by cst_idy", con, adOpenKeyset, adLockPessimistic
    If RSINL.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
        LVP.ColumnHeaders.Add , , "Party Name", 2000
        LVP.ColumnHeaders.Add , , "Party ID", 1500
        LVP.ColumnHeaders.Add , , "Market", 2000
      '  LVP.ColumnHeaders.Add , , "Act. Date", 1200
     '   LVP.ColumnHeaders.Add , , "Party Id", 0
        LVP.ListItems.CLEAR
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_nme,act_idy from act", con, adOpenKeyset
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset
        Dim count As Integer
        count = 0
        While Not RSINL.EOF
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_nme,act_idy from act where act_idy='" & Trim(RSINL(2)) & "'AND ACT_NME LIKE '" & txtcRetailer & "%'", con, adOpenKeyset
            If rsmar.State = 1 Then rsmar.Close
            rsmar.Open "SELECT mar.mar_idy,mar_nme,act_idy from mar,act where mar.mar_idy like act.mar_idy and act_idy='" & Trim(RSINL(2)) & "'", con, adOpenKeyset, adLockPessimistic
            
            If Not rsact.EOF Then
            count = count + 1
                Set ls = LVP.ListItems.Add(, , rsact(0))
                ls.ListSubItems.Add , , Trim(rsact(1)) & ""
                 If Not rsmar.EOF Then ls.ListSubItems.Add , , rsmar(1)
'            Else
'                Set ls = LVP.ListItems.Add(, , "")
            End If
'            ls.ListSubItems.Add , , Trim(RSINL(2)) & ""

'             If Not rsmar.EOF Then ls.ListSubItems.Add , , rsmar(1)
'            If IsNull(RSINL(0)) = True Then
'            ls.ListSubItems.Add , , " "
'            Else
'            ls.ListSubItems.Add , , RSINL(0)
'            End If
'            If rsprd.State = 1 Then rsprd.Close
'            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", CON, adOpenKeyset
'            If Not rsprd.EOF Then ls.ListSubItems.Add , , rsprd(0)
'            ls.ListSubItems.Add , , Trim(RSINL(1)) & ""
'            ls.ListSubItems.Add , , Trim(RSINL(2)) & ""
            RSINL.MoveNext
        Wend
        If count > 0 Then
            SSTab1.TabEnabled(3) = True
            LVP.Visible = True
            LVP.SetFocus
            SSTab1.Tab = 3
        Else
            MsgBox "Party not found", vbInformation + vbOKOnly, "DAS Version 4.0"
            SSTab1.Tab = 5
            txtcRetailer = ""
            txtcRetailer.SetFocus
            SSTab1.TabEnabled(3) = False
        End If
    Else
        MsgBox "Party ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
        SSTab1.Tab = 5
        txtcRetailer = ""
        txtcRetailer.SetFocus
        SSTab1.TabEnabled(3) = False
    End If
Else
            'retid = LVP.SelectedItem.ListSubItems(1).Text
            While Len(Trim(txtcRetailer)) < 10
            txtcRetailer = "0" & txtcRetailer
            Wend
            retid = txtcRetailer
            If rsact.State = 1 Then rsact.Close
            rsact.Open "SELECT ACT_IDY,ACT_NME FROM ACT WHER ACT_IDY='" & Trim(retid) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then txtcRetailer = rsact(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' and avt_sts='Yes' and (caf_sts is null or caf_sts<>'Yes' or caf_sts='') and cst_idy='" & retid & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSC.Rows = 1
                MSC.Cols = 7
                MSC.TextMatrix(0, 0) = "SNO"
                MSC.ColWidth(0) = 700
                MSC.TextMatrix(0, 1) = "Mobile ID"
                MSC.ColWidth(1) = 1300
                MSC.ColAlignment(1) = 1
                MSC.TextMatrix(0, 2) = "Invoice ID"
                MSC.ColWidth(2) = 1300
                MSC.TextMatrix(0, 3) = "Status"
                MSC.ColWidth(3) = 800
                MSC.TextMatrix(0, 4) = "SKU Name"
                MSC.ColWidth(4) = 2200
                MSC.TextMatrix(0, 5) = "Party"
                MSC.ColWidth(5) = 1740
                MSC.TextMatrix(0, 6) = "Act. Date"
                MSC.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSC.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSC.TextMatrix(r, 0) = r
                    MSC.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSC.TextMatrix(r, 2) = ""
                    Else
                    MSC.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSC.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSC.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSC.TextMatrix(r, 5) = Trim(rsact(1))
                    MSC.TextMatrix(r, 6) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 5
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                MSC.Col = 3
                MSC.SetFocus
             Else
                MsgBox "Party ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
                SSTab1.Tab = 5
                txtcRetailer.SetFocus
                SSTab1.TabEnabled(3) = False
                txtcRetailer = ""
                MSC.CLEAR
                MSC.Rows = 2
            End If
            End If
End If
End Sub

Private Sub txtcSalesman_GotFocus()
'code added by pkb
chkCAll.Enabled = False
     optSmn.Value = True
        txtcInvoice = ""
        txtcRetailer = ""
        txtcMobile = ""
        If txtcSalesman = "" Then
        MSC.CLEAR
        End If
End Sub

Private Sub txtcSalesman_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    txtcSalesman = ""
    optRet.Value = True
    txtcRetailer.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    txtcSalesman = ""
    'cafdate.SetFocus
    'optDate.Value = True
    optRet.Value = False
    If chkCAll.Enabled = False Then chkCAll.Enabled = True
    chkCAll.SetFocus
    chkCAll.Value = 1
ElseIf KeyCode = 39 And Shift = 1 Then
    txtcSalesman = ""
    If txtcInvoice.Enabled = False Then txtcInvoice.Enabled = True
    txtcInvoice.SetFocus
    optInvoice.Value = True
End If
If KeyCode = 13 Then
If Not Val(Trim(txtcSalesman)) > 0 Then
    Dim ls As ListItem
    Label49 = "Select Invoice (Salesman) for CAF"
    SSTab1.TabEnabled(3) = True
    If RSINL.State = 1 Then RSINL.Close
'    RSINL.Open "select sal_idy,avt_dat,smn_idy,prd_idy,txn_typ,cst_idy from avt where sal_typ like 'SAL' and avt_sts='Yes' and (caf_sts is null or caf_sts='' or caf_sts<>'Yes') group by smn_idy", CON, adOpenKeyset, adlockpessimistic
'    If RSINL.RecordCount > 0 Then
'        LVP.ColumnHeaders.CLEAR
'        LVP.ColumnHeaders.Add , , "Salesman"
'        LVP.ColumnHeaders.Add , , "Invoice ID"
'        LVP.ColumnHeaders.Add , , "Party"
'        LVP.ColumnHeaders.Add , , "SKU Name"
'        LVP.ColumnHeaders.Add , , "Act. Date"
'        LVP.ColumnHeaders.Add , , "smnid", 0
'        LVP.ListItems.CLEAR
'        If RSSMN.State = 1 Then RSSMN.Close
'        RSSMN.Open "select smn_nme,smn_idy from smn", CON, adOpenKeyset, adlockpessimistic
'        If rsact.State = 1 Then rsact.Close
'        rsact.Open "select act_nme,act_idy from act", CON, adOpenKeyset, adlockpessimistic
'        If rsprd.State = 1 Then rsprd.Close
'        rsprd.Open "select prd_nme,prd_idy from prd", CON, adOpenKeyset, adlockpessimistic
'        While Not RSINL.EOF
'            If RSSMN.RecordCount > 0 Then RSSMN.MoveFirst
'            RSSMN.Find "smn_idy='" & Trim(RSINL(2)) & "'"
'            If Not RSSMN.EOF Then Set ls = LVP.ListItems.Add(, , RSSMN(0))
'            If IsNull(RSINL(0)) = True Then
'            ls.ListSubItems.Add , , " "
'            Else
'            ls.ListSubItems.Add , , RSINL(0)
'            End If
''           If rsact.RecordCount > 0 Then rsact.MoveFirst
''           rsact.Find "act_idy='" & Trim(RSINL(5)) & "'"
'            If rsact.State = 1 Then rsact.Close
'            rsact.Open "select act_nme,act_idy from act where act_idy='" & Trim(RSINL(5)) & "'", CON, adOpenKeyset, adlockpessimistic
'            If Not rsact.EOF Then ls.ListSubItems.Add , , rsact(0)
''           If rsprd.RecordCount > 0 Then rsprd.MoveFirst
''           rsprd.Find "prd_idy='" & Trim(RSINL(3)) & "'"
'            If rsprd.State = 1 Then rsprd.Close
'            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", CON, adOpenKeyset, adlockpessimistic
'            If Not rsprd.EOF Then ls.ListSubItems.Add , , rsprd(0)
'            ls.ListSubItems.Add , , RSINL(1)
'            ls.ListSubItems.Add , , RSINL(2)
'            RSINL.MoveNext
'        Wend

 If RSSMN.State = 1 Then RSSMN.Close
'     RSSMN.Open "select smn_nme,smn_idy from smn", CON, adOpenKeyset, adlockpessimistic
 RSSMN.Open "select smn_nme,smn_idy from smn WHERE smn_nme like '" & Trim(txtcSalesman) & "%' ", con, adOpenKeyset, adLockPessimistic
     If RSSMN.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
       
        LVP.ColumnHeaders.Add , , "Salesman NAME"
        LVP.ColumnHeaders.Add , , "Salesman ID"
        LVP.ListItems.CLEAR
     End If
    If RSSMN.RecordCount > 0 Then
  RSSMN.MoveFirst
  While Not RSSMN.EOF
        Set ls = LVP.ListItems.Add(, , RSSMN(0))
        ls.ListSubItems.Add , , RSSMN(1)
        RSSMN.MoveNext
  Wend



        SSTab1.TabEnabled(3) = True
        LVP.Visible = True
        LVP.SetFocus
        SSTab1.Tab = 3
    Else
        MsgBox "Salesman not Found", vbInformation + vbOKOnly, "DAS Version 4.0"
        SSTab1.Tab = 5
        txtcSalesman = ""
        txtcSalesman.SetFocus
        SSTab1.TabEnabled(3) = False
    End If
    
Else
        While Len(Trim(txtcSalesman)) < 10
        txtcSalesman = "0" & txtcSalesman
        Wend
        
        If rsAvt.State = 1 Then rsAvt.Close
        'txtcSalesman = LVP.SelectedItem.Text
        'smnid = LVP.SelectedItem.ListSubItems(1).Text
        rsAvt.Open "select SMN_idy from aVt WHERE SMN_IDY = '" & Trim(txtcSalesman) & "'", con, adOpenKeyset, adLockPessimistic
        If rsAvt.RecordCount > 0 Then           'CONDITION ADDED BY PKB
        
            smnid = ""
            smnid = txtcSalesman
'            smnid = LVP.SelectedItem.ListSubItems(1).Text
        
       '     smnid = LVP.SelectedItem.ListSubItems(5).Text
            If RSSMN.State = 1 Then RSSMN.Close
            RSSMN.Open "SELECT SMN_IDY,SMN_NME FROM SMN WHERE SMN_IDY='" & Trim(smnid) & "'", con, adOpenKeyset, adLockPessimistic
            If Not RSSMN.EOF Then txtcSalesman = RSSMN(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts,smn_idy from avt where sal_typ='SAL' and avt_sts='Yes' and (caf_sts is null or caf_sts<>'Yes' or caf_sts='') and smn_idy='" & Trim(smnid) & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSC.Rows = 1
                MSC.Cols = 7
                MSC.TextMatrix(0, 0) = "SNO"
                MSC.ColWidth(0) = 1300
                MSC.TextMatrix(0, 1) = "Mobile ID"
                MSC.ColWidth(1) = 1300
                MSC.ColAlignment(1) = 1
                MSC.TextMatrix(0, 2) = "Invoice ID"
                MSC.ColWidth(2) = 1300
                MSC.TextMatrix(0, 3) = "Status"
                MSC.ColWidth(3) = 800
                MSC.TextMatrix(0, 4) = "SKU Name"
                MSC.ColWidth(4) = 2200
                MSC.TextMatrix(0, 5) = "Party"
                MSC.ColWidth(5) = 1740
                MSC.TextMatrix(0, 6) = "Act. Date"
                MSC.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSC.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSC.TextMatrix(r, 0) = r
                    MSC.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSC.TextMatrix(r, 2) = ""
                    Else
                    MSC.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSC.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSC.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSC.TextMatrix(r, 5) = Trim(rsact(1))
                    MSC.TextMatrix(r, 6) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 5
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                MSC.Col = 3
                MSC.SetFocus
                End If
            Else 'MODIFIED BY PK'B
            SSTab1.Tab = 5
            MsgBox "No records found", vbInformation + vbOKOnly, "DAS Version 4.0"
            'LVP.SetFocus
            txtcSalesman = ""
            txtcSalesman.SetFocus
            MSC.Rows = 1
'            If chkCAll.Enabled = False Then chkCAll.Enabled = True
'            chkCAll.SetFocus
            End If
    
    End If
End If
End Sub

Private Sub txtInvoice_GotFocus()
'code added by pkb
Option1(1).Enabled = True
Option1(1).Value = True
txtMobile = ""
txtRetailer = ""
txtSalesman = ""
If txtInvoice = "" Then
MSA.CLEAR
End If
End Sub

Private Sub txtInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    txtInvoice = ""
    Option1(0).Value = True
    txtMobile.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    txtInvoice = ""
    Option1(3).Value = True
    txtRetailer.SetFocus
ElseIf KeyCode = 37 And Shift = 1 Then
    txtSalesman.SetFocus
    Option1(4).Value = True
End If
If KeyCode = 13 Then
    Dim ls As ListItem
    Label49 = "Select Invoice for Mobile(s) Activation"
    SSTab1.TabEnabled(3) = True
'    If INVIDY = "Caf" Or MBLIDY = "Chr" Then
'        qry = "avt_sts='Yes'"
'    Else
'        qry = "(avt_sts is null or avt_sts='No' or avt_sts='')"
'    End If
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "select sal_idy,sal_dat,cst_idy,prd_idy,txn_typ from avt where SAL_IDY IS NOT NULL AND sal_typ like 'SAL' and (avt_sts is null or avt_sts='No' or avt_sts='') group BY sal_idy", con, adOpenKeyset, adLockPessimistic
        
    If IsNull(RSINL(0)) = True Then
        MsgBox "Cannot Activate from Sales Invoice!", vbInformation + vbOKOnly, "DAS Version 4.0"
        txtInvoice.SetFocus
    Else
    If RSINL.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
        LVP.ColumnHeaders.Add , , "Invoice ID"
        LVP.ColumnHeaders.Add , , "SKU Name", 2500
        LVP.ColumnHeaders.Add , , "Party", 2500
        LVP.ColumnHeaders.Add , , "Invoice Date", 1200
        LVP.ListItems.CLEAR
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_idy,prd_nme from prd", con, adOpenKeyset, adLockPessimistic
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
        While Not RSINL.EOF
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_idy,prd_nme from prd where prd_idy='" & RSINL(3) & "'", con, adOpenKeyset, adLockPessimistic
            If IsNull(RSINL(0)) = True Then
            Set ls = LVP.ListItems.Add(, , " ")
            Else
            Set ls = LVP.ListItems.Add(, , RSINL(0))
            End If
            If Not rsprd.EOF Then ls.ListSubItems.Add , , rsprd(1)
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act where act_idy='" & RSINL(2) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then ls.ListSubItems.Add , , rsact(1)
            If IsNull(RSINL(1)) = True Then
            ls.ListSubItems.Add , , " "
            Else
            ls.ListSubItems.Add , , RSINL(1)
            End If
            RSINL.MoveNext
        Wend
        SSTab1.TabEnabled(3) = True
        LVP.Visible = True
        LVP.SetFocus
        SSTab1.Tab = 3
    Else
        MsgBox "Invoice ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
'        If INVIDY = "Caf" Then
'            SSTab1.Tab = 5
'            txtcInvoice.SetFocus
'        ElseIf INVIDY = "Chr" Then
'            SSTab1.Tab = 6
'            txtrInvoice.SetFocus
'        Else
            SSTab1.Tab = 2
            txtInvoice.SetFocus
            SSTab1.TabEnabled(3) = False
'        End If
'        INVIDY = ""
    End If
End If
    'INVIDY = ""
End If
End Sub

Private Sub txtMobile_GotFocus()
'code added by pkb
Option1(0).Enabled = True
Option1(0).Value = True
txtInvoice = ""
txtRetailer = ""
txtSalesman = ""
If txtMobile = "" Then
MSA.CLEAR
End If
End Sub

Private Sub txtMobile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 And Shift = 1 Then
    txtMobile = ""
    Option1(1).Value = True
    txtInvoice.SetFocus
ElseIf KeyCode = 37 And Shift = 1 Then
txtRetailer.SetFocus
Option1(3).Value = True
End If
If KeyCode = 13 Then
If Not Len(Trim(txtMobile)) > 9 Then
    Dim ls As ListItem
    Label49 = "Select Mobile NO"
    SSTab1.TabEnabled(3) = True
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "select mbl_idy,sal_idy,sal_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' AND (avt_sts is null or avt_sts='No' or avt_sts='') and mbl_idy like '" & Trim(txtMobile) & "%' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
    If RSINL.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
        LVP.ColumnHeaders.Add , , "Mobile ID"
        LVP.ColumnHeaders.Add , , "Invoice ID"
        LVP.ColumnHeaders.Add , , "SKU Name"
        LVP.ColumnHeaders.Add , , "Party"
        LVP.ColumnHeaders.Add , , "Invoice Date"
        LVP.ListItems.CLEAR
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
        While Not RSINL.EOF
            Set ls = LVP.ListItems.Add(, , RSINL(0))
            If IsNull(RSINL(1)) = True Then
            ls.ListSubItems.Add , , " "
            Else
            ls.ListSubItems.Add , , RSINL(1)
            End If
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsprd.EOF Then ls.ListSubItems.Add , , Trim(rsprd(0))
            ls.ListSubItems.Add , , Trim(RSINL(4)) & ""
            If IsNull(RSINL(2)) = True Then
            ls.ListSubItems.Add , , " "
            Else
            ls.ListSubItems.Add , , RSINL(2)
            End If
            RSINL.MoveNext
        Wend
        LVP.Visible = True
        LVP.SetFocus
        SSTab1.Tab = 3
    Else
        MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
        SSTab1.Tab = 2
        txtMobile.SetFocus
    End If
Else
        If Len(Trim(txtMobile)) > 9 Then
        MBLID = Trim(txtMobile)
'MBLID = LVP.SelectedItem.Text
        If RSINL.State = 1 Then RSINL.Close
        Set RSINL = Nothing
        RSINL.Open "select mbl_idy,sal_idy,sal_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' AND (avt_sts is null or avt_sts='No' or avt_sts='') and mbl_idy='" & Trim(MBLID) & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
        If RSINL.RecordCount > 0 Then
            
            txtMobile.Text = MBLID
    '       MSA.FixedRows = 0
            MSA.Rows = 1
            MSA.Cols = 7
            MSA.TextMatrix(0, 0) = "SNO"
             MSA.ColWidth(0) = 700
            MSA.TextMatrix(0, 1) = "Mobile ID"
            MSA.ColWidth(1) = 1300
            MSA.ColAlignment(1) = 1
            MSA.TextMatrix(0, 2) = "Invoice ID"
            MSA.ColWidth(2) = 1300
            MSA.TextMatrix(0, 3) = "Status"
            MSA.ColWidth(3) = 800
            MSA.TextMatrix(0, 4) = "SKU Name"
            MSA.ColWidth(4) = 2200
            MSA.TextMatrix(0, 5) = "Party"
            MSA.ColWidth(5) = 1740
            MSA.TextMatrix(0, 6) = "Invoice Date"
            MSA.ColWidth(6) = 1200
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
            MSA.Rows = RSINL.RecordCount + 1
    '        MSA.FixedRows = 1
            r = 1
            While Not RSINL.EOF
                MSA.TextMatrix(r, 0) = r
                MSA.TextMatrix(r, 1) = Trim(RSINL(0))
                If IsNull(RSINL(1)) = True Then
                MSA.TextMatrix(r, 2) = " "
                Else
                MSA.TextMatrix(r, 2) = Trim(RSINL(1))
                End If
                MSA.TextMatrix(r, 3) = "No" 'Trim(RSINL(6)) & ""
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsprd.EOF Then MSA.TextMatrix(r, 4) = Trim(rsprd(0))
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsact.EOF Then MSA.TextMatrix(r, 5) = Trim(rsact(1))
                If IsNull(RSINL(2)) = True Then
                MSA.TextMatrix(r, 6) = " "
                Else
                MSA.TextMatrix(r, 6) = Trim(RSINL(2))
                End If
                r = r + 1
                RSINL.MoveNext
            Wend
            SSTab1.Tab = 2
            SSTab1.TabEnabled(3) = False
            avtdate.SetFocus
            'MSA.Col = 2
            'MSA.SetFocus
        Else
            MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
            SSTab1.Tab = 2
            'actDate.SetFocus
            'Option1(2).Value = True
            If chkAAll.Enabled = False Then chkAAll.Enabled = True
            chkAAll.SetFocus
            chkAAll.Value = 1
            SSTab1.TabEnabled(3) = False
            End If
        End If
End If
End If
'MBLIDY = ""
End Sub

Private Sub txtMobile_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub txtModInv_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   If SRT = "SAL" Then
        '-----------
        lblno(0).Visible = False
        lblno(3).Visible = False
        lblno(2).Visible = False
        lblno(1).Visible = False
        '-----------
        If chkModDate.Value = 1 Then DQRY = " and INL.INV_DAT=CTOD('" & modDate & "')"
        If rsMod.State = 1 Then rsMod.Close
        rsMod.Open "select inl.inv_idy,inl.cst_idy,inl.inv_dat from inl,inv,prd where inl.inv_idy like inv.inv_idy and inv.prd_idy like prd.prd_idy and prd.exc_inf in('y','Y') and cst_idy like '" & Trim(txtModRet) & "%'" & DQRY & " group by inl.inv_idy", con, adOpenStatic
        If Not rsMod.EOF Then
            LVP.ColumnHeaders.CLEAR
            LVP.ColumnHeaders.Add , , "Invoice ID", 2000
            LVP.ColumnHeaders.Add , , "Party Name", 2500
            LVP.ColumnHeaders.Add , , "Invoice Date", 1800
            LVP.ListItems.CLEAR
            While Not rsMod.EOF
                Set ls = LVP.ListItems.Add(, , Trim(rsMod(0)))
                ls.ListSubItems.Add , , Trim(rsMod(1))
                ls.ListSubItems.Add , , Trim(rsMod(2))
                rsMod.MoveNext
            Wend
            Label49 = "Select Invoice (Retail Circulation)"
            SSTab1.Tab = 3
            LVP.Visible = True
            LVP.SetFocus
        Else
            MsgBox "Invoice Not Found, Define Invoice !", vbExclamation, "DAS Version 4.0"
            LVMOD.ListItems.CLEAR
            SSTab1.Tab = 7
            txtModInv.Enabled = True
            txtModInv.SetFocus
        End If
    Else
        Label49 = "Select Invoice for Adjust Returns"
        invlist_RETURNS
    End If
End If
End Sub

Private Sub txtModRet_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
If SRT = "SAL" Then
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
            rs.Open " select * from act where act_idy like '" & Sincrement(txtModRet) & "' and grp_clf='Debtor' and cst_typ in ('Retailer','Key A/c','Local Key A/c','Regnl Key A/c')", con, adOpenStatic
            If rs.RecordCount > 0 Then
                txtModRet = Sincrement(txtRet)
                txtModRetName = Trim(rs("ACT_nme"))
                'txtRet.Enabled = False
                txtModInv.SetFocus
            Else
                actlistMod
            End If
ElseIf SRT = "OAP" Then
    If rs.State = 1 Then rs.Close
           rs.Open " select * from act where act_idy like '" & Sincrement(txtModRet) & "' and grp_clf='Creditor' and cst_typ in ('Retailer','Key A/c','Local Key A/c','Regnl Key A/c') and act_typ='PUR'", con, adOpenStatic
            If rs.RecordCount > 0 Then
                txtModRet = Sincrement(txtRet)
                txtModRetName = Trim(rs("ACT_nme"))
                txtModInv.SetFocus
                'txtRet.Enabled = False
            Else
                actlistMod1
            End If
End If

End If
End Sub

Private Sub txtRetailer_GotFocus()
'code added by pkb
Option1(3).Enabled = True
Option1(3).Value = True
txtInvoice = ""
txtMobile = ""
txtSalesman = ""
If txtRetailer = "" Then
MSA.CLEAR
End If
End Sub

Private Sub txtRetailer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    txtRetailer = ""
    txtInvoice.SetFocus
    Option1(1).Value = True
ElseIf KeyCode = 38 And Shift = 1 Then
    txtRetailer = ""
    Option1(4).Value = True
    txtSalesman.SetFocus
ElseIf KeyCode = 39 And Shift = 1 Then
txtMobile.SetFocus
Option1(0).Value = True
End If
If KeyCode = 13 Then
If Not Val(Trim(txtRetailer)) > o Then
    Dim ls As ListItem
    Label49 = "Select Party Account"
    SSTab1.TabEnabled(3) = True
'    If RETIDY = "Caf" Or RETIDY = "Chr" Then
'        qry = "avt_sts='Yes'"
'    Else
'        qry = "(avt_sts is null or avt_sts='No' or avt_sts='')"
'    End If


    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "select sal_idy,sal_dat,cst_idy,prd_idy,txn_typ from avt where sal_typ like 'SAL' and (avt_sts is null or avt_sts='No' or avt_sts='') group by cst_idy", con, adOpenKeyset, adLockPessimistic
    
    If RSINL.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
        LVP.ColumnHeaders.Add , , "Party Name", 2000
        LVP.ColumnHeaders.Add , , "Party Id", 1500
        LVP.ColumnHeaders.Add , , "Market", 2000
  '      LVP.ColumnHeaders.Add , , "Invoice Date", 1200
   '     LVP.ColumnHeaders.Add , , "Party Id", 0
  ' End If
        LVP.ListItems.CLEAR
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_nme,act_idy from act", con, adOpenKeyset
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset
        RSINL.MoveFirst
        Dim count As Integer
        count = 0
        While Not RSINL.EOF
            If rsact.State = 1 Then rsact.Close
               'RSSMN.Open "select smn_nme,smn_idy from smn WHERE smn_nme like '" & Trim(txtSalesman) & "%' ", CON, adOpenKeyset, adlockpessimistic
            rsact.Open "select act_nme,act_idy from act where act_idy='" & Trim(RSINL(2)) & "' AND ACT_NME LIKE '" & txtRetailer & "%'", con, adOpenKeyset
            If rsmar.State = 1 Then rsmar.Close
            rsmar.Open "SELECT mar.mar_idy,mar_nme,act_idy from mar,act where mar.mar_idy like act.mar_idy and act_idy='" & Trim(RSINL(2)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then
                count = count + 1
                Set ls = LVP.ListItems.Add(, , rsact(0))
                ls.ListSubItems.Add , , Trim(rsact(1)) & ""
        
            End If
            RSINL.MoveNext
        Wend
        If count > 0 Then
        LVP.SortOrder = lvwDescending
        SSTab1.TabEnabled(3) = True
        LVP.Visible = True
        LVP.SetFocus
        SSTab1.Tab = 3
      Else
        MsgBox "Party not found", vbInformation + vbOKOnly, "DAS Version 4.0"
         SSTab1.Tab = 2
         txtRetailer.SetFocus
         txtRetailer = ""
         SSTab1.TabEnabled(3) = False
    End If
'            If Not rsact.EOF Then
'             ls.ListSubItems.Add , , Trim(rsact(1)) & ""
'             End If
      '      If Not rsmar.EOF Then ls.ListSubItems.Add , , rsmar(1)
'            If IsNull(RSINL(0)) = True Then
'            ls.ListSubItems.Add , , " "
'            Else
'            ls.ListSubItems.Add , , RSINL(0)
'            End If
'            If rsprd.State = 1 Then rsprd.Close
'            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", CON, adOpenKeyset
'            If Not rsprd.EOF Then ls.ListSubItems.Add , , rsprd(0)
'            ls.ListSubItems.Add , , Trim(RSINL(1)) & ""
'            ls.ListSubItems.Add , , Trim(RSINL(2)) & ""
           
   
'        If RETIDY = "Caf" Then
'            SSTab1.Tab = 5
'            txtcRetailer.SetFocus
'        ElseIf RETIDY = "Chr" Then
'            SSTab1.Tab = 6
'            txtrRetailer.SetFocus
'        Else
'            SSTab1.Tab = 2
'            txtRetailer.SetFocus
'            SSTab1.TabEnabled(3) = False
''        End If
''        RETIDY = ""
    End If
    
Else
While Len(Trim(txtRetailer)) < 10
    txtRetailer = "0" & txtRetailer
Wend
retid = txtRetailer
'retid = LVP.SelectedItem.ListSubItems(1).Text
            If rsact.State = 1 Then rsact.Close
            rsact.Open "SELECT ACT_IDY,ACT_NME FROM ACT WHER ACT_IDY='" & Trim(txtRetailer) & "'", con, adOpenKeyset, adLockPessimistic
          '  If Not rsact.EOF Then txtRetailer = rsact(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,sal_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' and (avt_sts is null or avt_sts='No' or avt_sts='') and cst_idy='" & retid & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSA.Rows = 1
                MSA.Cols = 7
                MSA.TextMatrix(0, 0) = "Mobile ID"
                MSA.ColWidth(0) = 700
                MSA.TextMatrix(0, 1) = "Mobile ID"
                MSA.ColWidth(1) = 1300
                MSA.ColAlignment(1) = 1
                MSA.TextMatrix(0, 2) = "Invoice ID"
                MSA.ColWidth(2) = 1300
                MSA.TextMatrix(0, 3) = "Status"
                MSA.ColWidth(3) = 800
                MSA.TextMatrix(0, 4) = "SKU Name"
                MSA.ColWidth(4) = 2200
                MSA.TextMatrix(0, 5) = "Party"
                MSA.ColWidth(5) = 1740
                MSA.TextMatrix(0, 6) = "Invoice Date"
                MSA.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSA.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSA.TextMatrix(r, 0) = r
                    MSA.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSA.TextMatrix(r, 2) = " "
                    Else
                    MSA.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSA.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSA.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSA.TextMatrix(r, 5) = Trim(rsact(1))
                    If IsNull(rsAvt(2)) = True Then
                    MSA.TextMatrix(r, 6) = " "
                    Else
                    MSA.TextMatrix(r, 6) = Trim(rsAvt(2))
                    End If
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                txtRetailer = rsact(1)
                SSTab1.Tab = 2
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                avtdate.SetFocus
                'MSA.Col = 2
                'MSA.SetFocus
            Else
            MsgBox "Party ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
            txtRetailer = ""
            txtRetailer.SetFocus
            MSA.CLEAR
            End If
End If
End If
End Sub

Private Sub txtrInvoice_GotFocus()
'CODE ADDED BY PKB
optrInvoice.Value = True
        txtrSalesman = ""
        txtrRetailer = ""
        txtrMobile = ""
        If txtrInvoice = "" Then
            MSR.CLEAR
        End If
End Sub

Private Sub txtrInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    txtrInvoice = ""
    optrMobile.Value = True
    txtrMobile.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    txtrInvoice = ""
    optrRet.Value = True
    txtrRetailer.SetFocus
    
ElseIf KeyCode = 37 And Shift = 1 Then
    txtrInvoice = ""
    txtrSalesman.SetFocus
    optrSmn.Value = True
End If
If KeyCode = 13 Then
    Dim ls As ListItem
    Label49 = "Select Invoice for Mobile(s) Activation for Churn"
    SSTab1.TabEnabled(3) = True
   
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "select sal_idy,avt_dat,cst_idy,prd_idy,txn_typ,avt_sts,caf_sts from avt where SAL_IDY IS NOT NULL AND sal_typ like 'SAL' and avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') group BY sal_idy", con, adOpenKeyset, adLockPessimistic
    
    If IsNull(RSINL(0)) = True Then
    MsgBox "Can't Churn through Invoice!", vbInformation + vbOKOnly, "DAS Version 4.0"
    txtrInvoice.SetFocus
    Else
    If RSINL.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
        LVP.ColumnHeaders.Add , , "Invoice ID"
        LVP.ColumnHeaders.Add , , "SKU Name", 2500
        LVP.ColumnHeaders.Add , , "Party", 2500
        LVP.ColumnHeaders.Add , , "Act. Date", 1200
        LVP.ListItems.CLEAR
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_idy,prd_nme from prd", con, adOpenKeyset, adLockPessimistic
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
        While Not RSINL.EOF
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_idy,prd_nme from prd where prd_idy='" & RSINL(3) & "'", con, adOpenKeyset, adLockPessimistic
            If IsNull(RSINL(0)) = True Then
            Set ls = LVP.ListItems.Add(, , " ")
            Else
            Set ls = LVP.ListItems.Add(, , RSINL(0))
            End If
            If Not rsprd.EOF Then ls.ListSubItems.Add , , rsprd(1)
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act where act_idy='" & RSINL(2) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then ls.ListSubItems.Add , , rsact(1)
            ls.ListSubItems.Add , , RSINL(1)
            RSINL.MoveNext
        Wend
        SSTab1.TabEnabled(3) = True
        LVP.Visible = True
        LVP.SetFocus
        SSTab1.Tab = 3
    Else
        MsgBox "Invoice ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
        SSTab1.Tab = 6
        txtrInvoice.SetFocus
        SSTab1.TabEnabled(3) = False
    End If
    End If
End If
End Sub

Private Sub txtrMobile_GotFocus()
optrMobile.Value = True
        txtrInvoice = ""
        txtrRetailer = ""
        txtrSalesman = ""
        If txtrMobile = "" Then
            MSR.CLEAR
        End If
        
End Sub

Private Sub txtrMobile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 And Shift = 1 Then
    txtrMobile = ""
    optrInvoice.Value = True
    txtrInvoice.SetFocus
    
ElseIf KeyCode = 37 And Shift = 1 Then
    txtrMobile = ""
    txtrRetailer.SetFocus
    optrRet.Value = True
    
End If
If KeyCode = 13 Then
If Not Len(Trim(txtrMobile)) > 9 Then
    Dim ls As ListItem
    Label49 = "Select Mobile NO for Churn"
    SSTab1.TabEnabled(3) = True
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts,caf_sts from avt where sal_typ='SAL' AND avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') and mbl_idy like '" & Trim(txtrMobile) & "%'", con, adOpenKeyset, adLockPessimistic
    If RSINL.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
        LVP.ColumnHeaders.Add , , "Mobile ID"
        LVP.ColumnHeaders.Add , , "Invoice ID"
        LVP.ColumnHeaders.Add , , "SKU Name"
        LVP.ColumnHeaders.Add , , "Party"
        LVP.ColumnHeaders.Add , , "Act. Date"
        LVP.ListItems.CLEAR
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
        While Not RSINL.EOF
            Set ls = LVP.ListItems.Add(, , RSINL(0))
            If IsNull(RSINL(1)) = True Then
            ls.ListSubItems.Add , , " "
            Else
            ls.ListSubItems.Add , , RSINL(1)
            End If
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsprd.EOF Then ls.ListSubItems.Add , , Trim(rsprd(0))
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act where act_idy='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then ls.ListSubItems.Add , , Trim(rsact(1)) & ""
            ls.ListSubItems.Add , , RSINL(2)
            RSINL.MoveNext
        Wend
        LVP.Visible = True
        LVP.SetFocus
        SSTab1.Tab = 3
    Else
        MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
        SSTab1.Tab = 6
        txtrMobile.SetFocus
    End If
Else
            MBLID = Trim(txtrMobile)
            'MBLID = LVP.SelectedItem.Text
            txtrMobile.Text = MBLID
            If RSINL.State = 1 Then RSINL.Close
            Set RSINL = Nothing
                RSINL.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts,caf_sts from avt where sal_typ='SAL' AND avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') and mbl_idy='" & Trim(MBLID) & "'", con, adOpenKeyset, adLockPessimistic
                If RSINL.RecordCount > 0 Then
                    MSR.Rows = 1
                    MSR.Cols = 7
                    MSR.TextMatrix(0, 0) = "SNO"
                    MSR.ColWidth(0) = 700
                    MSR.TextMatrix(0, 1) = "Mobile ID"
                    MSR.ColWidth(1) = 1300
                    MSR.ColAlignment(1) = 1
                    MSR.TextMatrix(0, 2) = "Invoice ID"
                    MSR.ColWidth(2) = 1300
                    MSR.TextMatrix(0, 3) = "Status"
                    MSR.ColWidth(3) = 800
                    MSR.TextMatrix(0, 4) = "SKU Name"
                    MSR.ColWidth(4) = 2200
                    MSR.TextMatrix(0, 5) = "Party"
                    MSR.ColWidth(5) = 1740
                    MSR.TextMatrix(0, 6) = "Act. Date"
                    MSR.ColWidth(6) = 1200
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                    MSR.Rows = RSINL.RecordCount + 1
                    r = 1
                    While Not RSINL.EOF
                        MSR.TextMatrix(r, 0) = r
                        MSR.TextMatrix(r, 1) = Trim(RSINL(0))
                        If IsNull(RSINL(1)) = True Then
                        MSR.TextMatrix(r, 2) = " "
                        Else
                        MSR.TextMatrix(r, 2) = Trim(RSINL(1))
                        End If
                        MSR.TextMatrix(r, 3) = "No" 'Trim(RSINL(6)) & ""
                        If rsprd.State = 1 Then rsprd.Close
                        rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(RSINL(3)) & "'", con, adOpenKeyset, adLockPessimistic
                        If Not rsprd.EOF Then MSR.TextMatrix(r, 4) = Trim(rsprd(0))
                        If rsact.State = 1 Then rsact.Close
                        rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(RSINL(4)) & "'", con, adOpenKeyset, adLockPessimistic
                        If Not rsact.EOF Then MSR.TextMatrix(r, 5) = Trim(rsact(1))
                        MSR.TextMatrix(r, 6) = Trim(RSINL(2))
                        r = r + 1
                        RSINL.MoveNext
                    Wend
                    SSTab1.Tab = 6
                    SSTab1.TabEnabled(3) = False
                    avtcdate.SetFocus
                    'MSR.Col = 2
                    'MSR.SetFocus
                Else
                    MsgBox "Mobile ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
                    SSTab1.Tab = 6
                    txtrMobile.SetFocus
                    SSTab1.TabEnabled(3) = False
                End If
End If
End If
End Sub

Private Sub txtrMobile_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub txtrRetailer_GotFocus()
        optrRet.Value = True
        txtrInvoice = ""
        txtrSalesman = ""
        txtrMobile = ""
        If txtrRetailer = "" Then
        MSR.CLEAR
        End If
End Sub

Private Sub txtrRetailer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    txtrRetailer = ""
    optrInvoice.Value = True
    txtrInvoice.SetFocus
ElseIf KeyCode = 38 And Shift = 1 Then
    txtrRetailer = ""
    optrSmn.Value = True
    txtrSalesman.SetFocus
    
ElseIf KeyCode = 39 And Shift = 1 Then
    txtrRetailer = ""
    txtrMobile.SetFocus
    optrMobile.Value = True
    
End If
If KeyCode = 13 Then
If Not Val(Trim(txtrRetailer)) > 0 Then
    Dim ls As ListItem
    Label49 = "Select Party Account for Churn"
    SSTab1.TabEnabled(3) = True
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "select sal_idy,avt_dat,cst_idy,prd_idy,txn_typ from avt where sal_typ like 'SAL' and avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') group by cst_idy", con, adOpenKeyset, adLockPessimistic
    If RSINL.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
        LVP.ColumnHeaders.Add , , "Party Name", 2000
        LVP.ColumnHeaders.Add , , "Party ID", 1500
        LVP.ColumnHeaders.Add , , "Market", 2000
     '   LVP.ColumnHeaders.Add , , "Act. Date", 1200
     '   LVP.ColumnHeaders.Add , , "Party Id", 0
        LVP.ListItems.CLEAR
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_nme,act_idy from act", con, adOpenKeyset
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset
        Dim count As Integer
        count = 0
        'CODE ADDED BY PKB
            While Not RSINL.EOF
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_nme,act_idy from act where act_idy='" & Trim(RSINL(2)) & "'AND ACT_NME LIKE '" & txtrRetailer & "%'", con, adOpenKeyset
                If rsmar.State = 1 Then rsmar.Close
                rsmar.Open "SELECT mar.mar_idy,mar_nme,act_idy from mar,act where mar.mar_idy like act.mar_idy and act_idy='" & Trim(RSINL(2)) & "'", con, adOpenKeyset, adLockPessimistic
            
            If Not rsact.EOF Then
            count = count + 1
                Set ls = LVP.ListItems.Add(, , rsact(0))
                ls.ListSubItems.Add , , Trim(rsact(1)) & ""
                 If Not rsmar.EOF Then ls.ListSubItems.Add , , rsmar(1)
            End If
            RSINL.MoveNext
            Wend
                If count > 0 Then
                SSTab1.TabEnabled(3) = True
                LVP.Visible = True
                LVP.SetFocus
                SSTab1.Tab = 3
            Else
                MsgBox "Party not found", vbInformation + vbOKOnly, "DAS Version 4.0"
                SSTab1.Tab = 6
                txtrRetailer = ""
                txtrRetailer.SetFocus
                SSTab1.TabEnabled(3) = False
            End If
            Else
                MsgBox "Party ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
                SSTab1.Tab = 6
                txtrRetailer = ""
                txtrRetailer.SetFocus
                SSTab1.TabEnabled(3) = False
            End If
     '   While Not RSINL.EOF
'         If rsmar.State = 1 Then rsmar.Close
'            rsmar.Open "SELECT ACT.ACT_NME,mar_nme,act_idy,ACT_NME from mar,act where mar.mar_idy like act.mar_idy and act_idy='" & Trim(RSINL(2)) & "'AND ACT.ACT_NME LIKE '" & Trim(txtrRetailer) & "%' ", CON, adOpenKeyset, adlockpessimistic
'            txtcRetailer = rsmar.RecordCount
''            If rsact.State = 1 Then rsact.Close
''            rsact.Open "select act_nme,act_idy from act where act_idy='" & Trim(RSINL(2)) & "'", CON, adOpenKeyset
'            If rsmar.State = 1 Then rsmar.Close
'            rsmar.Open "SELECT mar.mar_idy,mar_nme,act_idy from mar,act where mar.mar_idy like act.mar_idy and act_idy='" & Trim(RSINL(2)) & "'", CON, adOpenKeyset, adlockpessimistic
'            If Not rsact.EOF Then
'                Set ls = LVP.ListItems.Add(, , rsact(0))
''            Else
''                Set ls = LVP.ListItems.Add(, , "")
'            End If
'             ls.ListSubItems.Add , , Trim(rsact(1)) & ""
'             If Not rsmar.EOF Then ls.ListSubItems.Add , , rsmar(1)
'            If IsNull(RSINL(0)) = True Then
'            ls.ListSubItems.Add , , " "
'            Else
'            ls.ListSubItems.Add , , RSINL(0)
'            End If
'            If rsprd.State = 1 Then rsprd.Close
'            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", CON, adOpenKeyset
'            If Not rsprd.EOF Then ls.ListSubItems.Add , , rsprd(0)
'            ls.ListSubItems.Add , , Trim(RSINL(1)) & ""
'            ls.ListSubItems.Add , , Trim(RSINL(2)) & ""
            
Else
While Len(Trim(txtrRetailer)) < 10
    txtrRetailer = "0" & txtrRetailer
Wend
            retid = Trim(txtrRetailer)
            'retid = LVP.SelectedItem.ListSubItems(1).Text
            If rsact.State = 1 Then rsact.Close
            rsact.Open "SELECT ACT_IDY,ACT_NME FROM ACT WHER ACT_IDY='" & Trim(retid) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsact.EOF Then txtrRetailer = rsact(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts from avt where sal_typ='SAL' and avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') and cst_idy='" & retid & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSR.Rows = 1
                MSR.Cols = 7
                MSR.TextMatrix(0, 0) = "SNO"
                MSR.ColWidth(0) = 700
                MSR.TextMatrix(0, 1) = "Mobile ID"
                MSR.ColWidth(1) = 1300
                MSR.ColAlignment(1) = 1
                MSR.TextMatrix(0, 2) = "Invoice ID"
                MSR.ColWidth(2) = 1300
                MSR.TextMatrix(0, 3) = "Status"
                MSR.ColWidth(3) = 800
                MSR.TextMatrix(0, 4) = "SKU Name"
                MSR.ColWidth(4) = 2200
                MSR.TextMatrix(0, 5) = "Party"
                MSR.ColWidth(5) = 1740
                MSR.TextMatrix(0, 6) = "Act. Date"
                MSR.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSR.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSR.TextMatrix(r, 0) = r
                    MSR.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSR.TextMatrix(r, 2) = " "
                    Else
                    MSR.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSR.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSR.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSR.TextMatrix(r, 5) = Trim(rsact(1))
                    MSR.TextMatrix(r, 6) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 6
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                avtcdate.SetFocus
            Else
                MsgBox "Party ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
                SSTab1.Tab = 6
                txtrRetailer.SetFocus
                SSTab1.TabEnabled(3) = False
                txtrRetailer = ""
                MSR.CLEAR
                MSR.Rows = 1
            End If
    End If
End If
End Sub

Private Sub txtrSalesman_GotFocus()
chkRAll.Enabled = False
        optrSmn.Value = True
        txtrInvoice = ""
        txtrRetailer = ""
        txtrMobile = ""
        If txtrSalesman = "" Then
            MSR.CLEAR
        End If
        
End Sub

Private Sub txtrSalesman_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    txtrSalesman = ""
    optrRet.Value = True
    txtrRetailer.SetFocus
    
ElseIf KeyCode = 39 And Shift = 1 Then
    txtrSalesman = ""
    txtrInvoice.SetFocus
    optrInvoice.Value = True
ElseIf KeyCode = 38 And Shift = 1 Then
    txtrSalesman = ""
    'ChrDate.SetFocus
    'optrDate.Value = True
    If chkRAll.Enabled = False Then chkRAll.Enabled = True
    chkRAll.SetFocus
    chkRAll.Value = 1
End If
If KeyCode = 13 Then
If Not Val(Trim(txtrSalesman)) > 0 Then
    Dim ls As ListItem
    Label49 = "Select Invoice (Salesman) for Churn"
    SSTab1.TabEnabled(3) = True
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "select sal_idy,avt_dat,smn_idy,prd_idy,txn_typ,cst_idy from avt where sal_typ like 'SAL' and avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') group by smn_idy", con, adOpenKeyset, adLockPessimistic
    If RSINL.RecordCount > 0 Then
    'CODE COMMENTED BY PKB
'        LVP.ColumnHeaders.CLEAR
'        LVP.ColumnHeaders.Add , , "Salesman"
'        LVP.ColumnHeaders.Add , , "Invoice ID"
'        LVP.ColumnHeaders.Add , , "Party"
'        LVP.ColumnHeaders.Add , , "SKU Name"
'        LVP.ColumnHeaders.Add , , "Act. Date"
'        LVP.ColumnHeaders.Add , , "smnid", 0
'        LVP.ListItems.CLEAR
        If RSSMN.State = 1 Then RSSMN.Close
        RSSMN.Open "select smn_nme,smn_idy from smn", con, adOpenKeyset, adLockPessimistic
        If rsact.State = 1 Then rsact.Close
        rsact.Open "select act_nme,act_idy from act", con, adOpenKeyset, adLockPessimistic
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
        'CODE ADDED BY PKB
         If RSSMN.State = 1 Then RSSMN.Close
'     RSSMN.Open "select smn_nme,smn_idy from smn", CON, adOpenKeyset, adlockpessimistic
     RSSMN.Open "select smn_nme,smn_idy from smn WHERE smn_nme like '" & Trim(txtrSalesman) & "%' ", con, adOpenKeyset, adLockPessimistic
     If RSSMN.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
       
        LVP.ColumnHeaders.Add , , "Salesman NAME"
        LVP.ColumnHeaders.Add , , "Salesman ID"
        LVP.ListItems.CLEAR
     End If
    If RSSMN.RecordCount > 0 Then
  RSSMN.MoveFirst
  While Not RSSMN.EOF
        Set ls = LVP.ListItems.Add(, , RSSMN(0))
        ls.ListSubItems.Add , , RSSMN(1)
        RSSMN.MoveNext
  Wend
        SSTab1.TabEnabled(3) = True
        LVP.Visible = True
        LVP.SetFocus
        SSTab1.Tab = 3
  Else
        MsgBox "Salesman not found", vbInformation + vbOKOnly, "DAS Version 4.0"
        SSTab1.Tab = 6
        txtrSalesman = ""
        txtrSalesman.SetFocus
        SSTab1.TabEnabled(3) = False
    End If
    'CODE COMMENTED BY PKB
    
'        While Not RSINL.EOF
'            If RSSMN.RecordCount > 0 Then RSSMN.MoveFirst
'            RSSMN.Find "smn_idy='" & Trim(RSINL(2)) & "'"
'            If Not RSSMN.EOF Then Set ls = LVP.ListItems.Add(, , RSSMN(0))
'            If IsNull(RSINL(0)) = True Then
'            ls.ListSubItems.Add , , " "
'            Else
'            ls.ListSubItems.Add , , RSINL(0)
'            End If
'            If rsact.State = 1 Then rsact.Close
'            rsact.Open "select act_nme,act_idy from act where act_idy='" & Trim(RSINL(5)) & "'", CON, adOpenKeyset, adlockpessimistic
'            If Not rsact.EOF Then ls.ListSubItems.Add , , rsact(0)
'            If rsprd.State = 1 Then rsprd.Close
'            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", CON, adOpenKeyset, adlockpessimistic
'            If Not rsprd.EOF Then ls.ListSubItems.Add , , rsprd(0)
'            ls.ListSubItems.Add , , RSINL(1)
'            ls.ListSubItems.Add , , RSINL(2)
'            RSINL.MoveNext
'        Wend
       
    Else
        MsgBox "Salesman not found", vbInformation + vbOKOnly, "DAS Version 4.0"
        SSTab1.Tab = 6
        txtrSalesman = ""
        txtrSalesman.SetFocus
        SSTab1.TabEnabled(3) = False
    End If
Else
        While Len(Trim(txtrSalesman)) < 10
        txtrSalesman = "0" & txtrSalesman
        Wend
        smnid = txtrSalesman

        'MODIFIED BY PKB
         If rsAvt.State = 1 Then rsAvt.Close
           ' txtrSalesman = LVP.SelectedItem.Text
           ' smnid = LVP.SelectedItem.ListSubItems(1).Text
         
            rsAvt.Open "select SMN_idy from aVt WHERE SMN_IDY = '" & Trim(smnid) & "'", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then

            smnid = LVP.SelectedItem.ListSubItems(1).Text   ' INSPITE OF 5 PLACED 1 IN LSI
            If RSSMN.State = 1 Then RSSMN.Close
            RSSMN.Open "SELECT SMN_IDY,SMN_NME FROM SMN WHERE SMN_IDY='" & Trim(smnid) & "'", con, adOpenKeyset, adLockPessimistic
            If Not RSSMN.EOF Then txtrSalesman = RSSMN(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,avt_dat,prd_idy,cst_idy,txn_typ,avt_sts,smn_idy from avt where sal_typ='SAL' and avt_sts='Yes' and (chr_sts is null or chr_sts<>'Yes' or chr_sts='') and smn_idy='" & Trim(smnid) & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSR.Rows = 1
                MSR.Cols = 7
                MSR.TextMatrix(0, 0) = "SNO"
                MSR.ColWidth(0) = 700
                MSR.TextMatrix(0, 1) = "Mobile ID"
                MSR.ColWidth(1) = 1300
                MSR.ColAlignment(1) = 1
                MSR.TextMatrix(0, 2) = "Invoice ID"
                MSR.ColWidth(2) = 1300
                MSR.TextMatrix(0, 3) = "Status"
                MSR.ColWidth(3) = 800
                MSR.TextMatrix(0, 4) = "SKU Name"
                MSR.ColWidth(4) = 2200
                MSR.TextMatrix(0, 5) = "Party"
                MSR.ColWidth(5) = 1740
                MSR.TextMatrix(0, 6) = "Act. Date"
                MSR.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSR.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSR.TextMatrix(r, 0) = r
                    MSR.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSR.TextMatrix(r, 2) = " "
                    Else
                    MSR.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSR.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSR.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSR.TextMatrix(r, 5) = Trim(rsact(1))
                    MSR.TextMatrix(r, 6) = Trim(rsAvt(2))
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 6
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                avtcdate.SetFocus
                End If
               Else
                SSTab1.Tab = 6
                MsgBox "No records found", vbInformation + vbOKOnly, "DAS Version 4.0"
                txtrSalesman = ""
                txtrSalesman.SetFocus
'                If chkRAll.Enabled = False Then chkRAll.Enabled = True
'                chkRAll.Value = 1
'                optrSmn.Value = False
'                chkRAll.SetFocus
End If
End If
End If
End Sub

Private Sub txtSalesman_gotfocus()
'code added by pkb
chkAAll.Value = 0
        chkAAll.Enabled = False
        Option1(4).Enabled = True
        Option1(4).Value = True
        txtInvoice = ""
        txtRetailer = ""
        txtMobile = ""
       If txtSalesman = "" Then
MSA.CLEAR
End If
End Sub

Private Sub txtSalesman_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 40 Then
    txtSalesman = ""
    txtRetailer.SetFocus
    Option1(3).Value = True
ElseIf KeyCode = 39 And Shift = 1 Then
    txtInvoice.SetFocus
    Option1(1).Value = True
ElseIf KeyCode = 38 And Shift = 1 Then
    Option1(4).Value = False
    If chkAAll.Enabled = False Then chkAAll.Enabled = True
    chkAAll.SetFocus
    chkAAll.Value = 1
    Else
End If
'
If KeyCode = 13 Then
 If Not Val(Trim(txtSalesman)) > 0 Then
    Dim ls As ListItem
    Label49 = "Select Invoice (Salesman)"
    SSTab1.TabEnabled(3) = True
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "select sal_idy,sal_dat,smn_idy,prd_idy,txn_typ,cst_idy from avt where  (avt_sts is null or avt_sts='No' or avt_sts='') group by smn_idy ", con, adOpenKeyset, adLockPessimistic
   ' RSINL.Open "select sal_idy,sal_dat,smn_idy,prd_idy,txn_typ,cst_idy from avt where sal_typ like 'SAL' and (avt_sts is null or avt_sts='No' or avt_sts='') group by smn_idy ", CON, adOpenKeyset, adlockpessimistic
    'RSINL.Open "select sal_idy,sal_dat,smn_idy,prd_idy,txn_typ,cst_idy from avt where (avt_sts is null or avt_sts='No' or avt_sts='') ", CON, adOpenKeyset, adlockpessimistic
     If RSSMN.State = 1 Then RSSMN.Close
     RSSMN.Open "select smn_nme,smn_idy from smn WHERE smn_nme like '" & Trim(txtSalesman) & "%' ", con, adOpenKeyset, adLockPessimistic
    ' If RSSMN.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
       
        LVP.ColumnHeaders.Add , , "Salesman NAME"
        LVP.ColumnHeaders.Add , , "Salesman ID"
        LVP.ListItems.CLEAR
    ' End If
    If RSSMN.RecordCount > 0 Then
  RSSMN.MoveFirst
  While Not RSSMN.EOF
        Set ls = LVP.ListItems.Add(, , RSSMN(0))
        ls.ListSubItems.Add , , RSSMN(1)
        RSSMN.MoveNext
  Wend
        SSTab1.TabEnabled(3) = True
        LVP.Visible = True
        LVP.SetFocus
        SSTab1.Tab = 3
  Else
        MsgBox "Salesman not found", vbInformation + vbOKOnly, "DAS Version 4.0"
        SSTab1.Tab = 2
        txtSalesman = ""
        txtSalesman.SetFocus
        SSTab1.TabEnabled(3) = False
 End If
'
'        SSTab1.TabEnabled(3) = True
'        LVP.Visible = True
'        LVP.SetFocus
'        SSTab1.Tab = 3
'End If
Else

 If rsAvt.State = 1 Then rsAvt.Close
'        txtSalesman = LVP.SelectedItem.Text
       ' smnid = LVP.SelectedItem.ListSubItems(1).Text
       smnid = txtSalesman
        smnid = txtSalesman
            While Len(Trim(smnid)) < 10
            smnid = "0" & smnid
            Wend
        rsAvt.Open "select SMN_idy from aVt WHERE SMN_IDY = '" & Trim(smnid) & "'", con, adOpenKeyset, adLockPessimistic
        If rsAvt.RecordCount > 0 Then           'MODIFIED BY PKB
        
            smnid = ""
'            smnid = LVP.SelectedItem.ListSubItems(1).Text
            smnid = txtSalesman
            While Len(Trim(smnid)) < 10
            smnid = "0" & smnid
            Wend
            
            If RSSMN.State = 1 Then RSSMN.Close
            RSSMN.Open "SELECT SMN_IDY,SMN_NME FROM SMN WHERE SMN_IDY='" & Trim(smnid) & "'", con, adOpenKeyset, adLockPessimistic
            If Not RSSMN.EOF Then txtSalesman = RSSMN(1)
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            rsAvt.Open "select mbl_idy,sal_idy,sal_dat,prd_idy,cst_idy,txn_typ,avt_sts,smn_idy from avt where sal_typ='SAL' and (avt_sts is null or avt_sts='No' or avt_sts='') and smn_idy='" & Trim(smnid) & "' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            If rsAvt.RecordCount > 0 Then
                MSA.Rows = 1
                MSA.Cols = 7
                MSA.TextMatrix(0, 0) = "SNO"
                MSA.ColWidth(0) = 700
                MSA.TextMatrix(0, 1) = "Mobile ID"
                MSA.ColWidth(1) = 1300
                MSA.ColAlignment(1) = 1
                MSA.TextMatrix(0, 2) = "Invoice ID"
                MSA.ColWidth(2) = 1300
                MSA.TextMatrix(0, 3) = "Status"
                MSA.ColWidth(3) = 800
                MSA.TextMatrix(0, 4) = "SKU Name"
                MSA.ColWidth(4) = 2200
                MSA.TextMatrix(0, 5) = "Party"
                MSA.ColWidth(5) = 1740
                MSA.TextMatrix(0, 6) = "Invoice Date"
                MSA.ColWidth(6) = 1200
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_nme,prd_idy from prd", con, adOpenKeyset, adLockPessimistic
                If rsact.State = 1 Then rsact.Close
                rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
                MSA.Rows = rsAvt.RecordCount + 1
                r = 1
                While Not rsAvt.EOF
                    MSA.TextMatrix(r, 0) = r
                    MSA.TextMatrix(r, 1) = Trim(rsAvt(0))
                    If IsNull(rsAvt(1)) = True Then
                    MSA.TextMatrix(r, 2) = " "
                    Else
                    MSA.TextMatrix(r, 2) = Trim(rsAvt(1))
                    End If
                    MSA.TextMatrix(r, 3) = "No" 'Trim(rsavt(6)) & ""
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_nme,prd_idy from prd WHERE prd_idy='" & Trim(rsAvt(3)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsprd.EOF Then MSA.TextMatrix(r, 4) = Trim(rsprd(0))
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(rsAvt(4)) & "'", con, adOpenKeyset, adLockPessimistic
                    If Not rsact.EOF Then MSA.TextMatrix(r, 5) = Trim(rsact(1))
                    If IsNull(rsAvt(2)) = True Then
                    MSA.TextMatrix(r, 6) = " "
                    Else
                    MSA.TextMatrix(r, 6) = Trim(rsAvt(2))
                    End If
                    r = r + 1
                    rsAvt.MoveNext
                Wend
                SSTab1.Tab = 2
                LVP.Visible = False
                SSTab1.TabEnabled(3) = False
                MSA.Col = 3
                MSA.SetFocus
                avtdate.SetFocus
            
            End If
            Else                        'MODIFIED BY PKB
            SSTab1.Tab = 2
            MsgBox "No records found", vbInformation + vbOKOnly, "DAS Version 4.0"
            txtSalesman = ""
            MSA.Rows = 1
'            chkAAll.SetFocus
            'LVP.SetFocus       'MODIFIED BY PKB
            End If               'MODIFIED BY PK'B
            End If
            
'If KeyCode = 13 Then
    'Dim ls As ListItem
'    Label49 = "Select Invoice (Salesman)"
'    SSTab1.TabEnabled(3) = True
'    If RSINL.State = 1 Then RSINL.Close
'    RSINL.Open "select sal_idy,sal_dat,smn_idy,prd_idy,txn_typ,cst_idy from avt where  (avt_sts is null or avt_sts='No' or avt_sts='') group by smn_idy ", CON, adOpenKeyset, adlockpessimistic
'   ' RSINL.Open "select sal_idy,sal_dat,smn_idy,prd_idy,txn_typ,cst_idy from avt where sal_typ like 'SAL' and (avt_sts is null or avt_sts='No' or avt_sts='') group by smn_idy ", CON, adOpenKeyset, adlockpessimistic
'    'RSINL.Open "select sal_idy,sal_dat,smn_idy,prd_idy,txn_typ,cst_idy from avt where (avt_sts is null or avt_sts='No' or avt_sts='') ", CON, adOpenKeyset, adlockpessimistic
'    If RSINL.RecordCount > 0 Then
'        LVP.ColumnHeaders.CLEAR
'        LVP.ColumnHeaders.Add , , "Salesman"
'        LVP.ColumnHeaders.Add , , "Invoice ID"
'        LVP.ColumnHeaders.Add , , "Party"
'        LVP.ColumnHeaders.Add , , "SKU Name"
'        LVP.ColumnHeaders.Add , , "Invoice Date"
'        LVP.ColumnHeaders.Add , , "Salesman ID", 0
'        LVP.ListItems.CLEAR
'        If RSSMN.State = 1 Then RSSMN.Close
'        RSSMN.Open "select smn_nme,smn_idy from smn", CON, adOpenKeyset, adlockpessimistic
'        If rsact.State = 1 Then rsact.Close
'        rsact.Open "select act_nme,act_idy from act", CON, adOpenKeyset, adlockpessimistic
'        If rsprd.State = 1 Then rsprd.Close
'        rsprd.Open "select prd_nme,prd_idy from prd", CON, adOpenKeyset, adlockpessimistic
   '     While Not RSSMN.EOF
           ' If RSSMN.RecordCount > 0 Then RSSMN.MoveFirst
          '  RSSMN.Find "smn_idy='" & Trim(RSINL(2)) & "'"
            'If Not RSSMN.EOF Then
'            LVP.ListItems.Add (RSSMN(1))
            'LVP.ListItems.Add (RSSMN(0))
'            Set ls = LVP.ListItems.Add(, , " ")
'            End If
'
'            If IsNull(RSINL(0)) = True Then
'            ls.ListSubItems.Add , , " "
'            Else
'            ls.ListSubItems.Add , , RSSMN(1)
'            End If
'            If rsact.State = 1 Then rsact.Close
'            rsact.Open "select act_nme,act_idy from act where act_idy='" & Trim(RSINL(5)) & "'", CON, adOpenKeyset, adlockpessimistic
'            If Not rsact.EOF Then ls.ListSubItems.Add , , Trim(rsact(0))
'            If rsprd.State = 1 Then rsprd.Close
'            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", CON, adOpenKeyset, adlockpessimistic
'            If Not rsprd.EOF Then ls.ListSubItems.Add , , rsprd(0)
'
'            If IsNull(RSINL(1)) = True Then
'            ls.ListSubItems.Add , , " "
'            Else
'            ls.ListSubItems.Add , , RSINL(1)
'            End If
'
'            If IsNull(RSINL(2)) = True Then
'            ls.ListSubItems.Add , , " "
'            Else
'            ls.ListSubItems.Add , , RSSMN(0)
'            End If

  '          RSSMN.MoveNext
     '   Wend
'        SSTab1.TabEnabled(3) = True
'        LVP.Visible = True
'        LVP.SetFocus
'        SSTab1.Tab = 3
'    ElseIf KeyCode = 27 Then
'    SSTab1.Tab = 0
'    Else
'        MsgBox "Salesman ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
'        SSTab1.Tab = 2
'        txtSalesman.SetFocus
'        SSTab1.TabEnabled(3) = False
   ' End If
'End If
'
'If KeyCode = 13 Then
'    Dim ls As ListItem
'    Label49 = "Select Invoice (Salesman)"
'    SSTab1.TabEnabled(3) = True
'    If RSINL.State = 1 Then RSINL.Close
'    RSINL.Open "select sal_idy,sal_dat,smn_idy,prd_idy,txn_typ,cst_idy from avt where  (avt_sts is null or avt_sts='No' or avt_sts='') group by smn_idy ", CON, adOpenKeyset, adlockpessimistic
'   ' RSINL.Open "select sal_idy,sal_dat,smn_idy,prd_idy,txn_typ,cst_idy from avt where sal_typ like 'SAL' and (avt_sts is null or avt_sts='No' or avt_sts='') group by smn_idy ", CON, adOpenKeyset, adlockpessimistic
'    'RSINL.Open "select sal_idy,sal_dat,smn_idy,prd_idy,txn_typ,cst_idy from avt where (avt_sts is null or avt_sts='No' or avt_sts='') ", CON, adOpenKeyset, adlockpessimistic
'    If RSINL.RecordCount > 0 Then
'        LVP.ColumnHeaders.CLEAR
'        LVP.ColumnHeaders.Add , , "Salesman"
'        LVP.ColumnHeaders.Add , , "Invoice ID"
'        LVP.ColumnHeaders.Add , , "Party"
'        LVP.ColumnHeaders.Add , , "SKU Name"
'        LVP.ColumnHeaders.Add , , "Invoice Date"
'        LVP.ColumnHeaders.Add , , "Salesman ID", 0
'        LVP.ListItems.CLEAR
'        If RSSMN.State = 1 Then RSSMN.Close
'        RSSMN.Open "select smn_nme,smn_idy from smn", CON, adOpenKeyset, adlockpessimistic
'        If rsact.State = 1 Then rsact.Close
'        rsact.Open "select act_nme,act_idy from act", CON, adOpenKeyset, adlockpessimistic
'        If rsprd.State = 1 Then rsprd.Close
'        rsprd.Open "select prd_nme,prd_idy from prd", CON, adOpenKeyset, adlockpessimistic
'        While Not RSINL.EOF
'            If RSSMN.RecordCount > 0 Then RSSMN.MoveFirst
'            RSSMN.Find "smn_idy='" & Trim(RSINL(2)) & "'"
'            If Not RSSMN.EOF Then
'            Set ls = LVP.ListItems.Add(, , RSSMN(0))
'            Else
'            Set ls = LVP.ListItems.Add(, , " ")
'            End If
'
'            If IsNull(RSINL(0)) = True Then
'            ls.ListSubItems.Add , , " "
'            Else
'            ls.ListSubItems.Add , , RSINL(0)
'            End If
'            If rsact.State = 1 Then rsact.Close
'            rsact.Open "select act_nme,act_idy from act where act_idy='" & Trim(RSINL(5)) & "'", CON, adOpenKeyset, adlockpessimistic
'            If Not rsact.EOF Then ls.ListSubItems.Add , , Trim(rsact(0))
'            If rsprd.State = 1 Then rsprd.Close
'            rsprd.Open "select prd_nme,prd_idy from prd where prd_idy='" & Trim(RSINL(3)) & "'", CON, adOpenKeyset, adlockpessimistic
'            If Not rsprd.EOF Then ls.ListSubItems.Add , , rsprd(0)
'
'            If IsNull(RSINL(1)) = True Then
'            ls.ListSubItems.Add , , " "
'            Else
'            ls.ListSubItems.Add , , RSINL(1)
'            End If
'
'            If IsNull(RSINL(2)) = True Then
'            ls.ListSubItems.Add , , " "
'            Else
'            ls.ListSubItems.Add , , RSINL(2)
'            End If
'
'            RSINL.MoveNext
'        Wend
'        SSTab1.TabEnabled(3) = True
'        LVP.Visible = True
'        LVP.SetFocus
'        SSTab1.Tab = 3
'    Else
'        MsgBox "Salesman ID not found", vbInformation + vbOKOnly, "DAS Version 4.0"
'        SSTab1.Tab = 2
'        txtSalesman.SetFocus
'        SSTab1.TabEnabled(3) = False
'    End If
'End If
End If

End Sub

Private Sub txtSalInv_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

On Error GoTo SAL_ERR
'For S = 0 To 5
'    If Ptno(S).Visible = True Then Ptno(S).Visible = False
'Next
If Trim(txtSalInv.Text) <> "" Then
        If RSINV.State = 1 Then RSINV.Close
        If Mid(Trim(txtSalInv.Text), 1, 3) = "SAL" Then
            RSINV.Open "select inv.inv_idy,inv.inv_dat,inv.prd_idy,SUM(inv.prd_qty),inl.cst_idy,inv.txn_typ,inl.smn_idy,exc_inf from inv,inl,prd where inv.inv_idy like inl.inv_idy and inv.inv_dat=inl.inv_dat AND inv.txn_typ in('SAL','STO','PRT') and inv.prd_idy like prd.prd_idy and prd.exc_inf='Y' and inv.stk_typ in('Salable','Free') and INL.inv_idy='" & Trim(txtSalInv.Text) & "'", con, adOpenKeyset, adLockPessimistic
        Else
            RSINV.Open "select inv.inv_idy,inv.inv_dat,inv.prd_idy,SUM(inv.prd_qty),inl.cst_idy,inv.txn_typ,inl.smn_idy,exc_inf from inv,inl,prd where inv.inv_idy like inl.inv_idy and inv.inv_dat=inl.inv_dat AND inv.txn_typ in('SAL','STO','PRT') and inv.prd_idy like prd.prd_idy and prd.exc_inf='Y' and inv.stk_typ in('Salable','Free') and INL.inv_idy='" & Trim(txtSalInv.Text) & "'", con, adOpenKeyset, adLockPessimistic
        End If
        If RSINV.EOF = True Then
            MsgBox "Invoice not Found!", vbOKOnly, "DAS Version 4.0"
            txtSalInv.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
            Dim pct_rst As New ADODB.Recordset
            Dim pct_rst1 As New ADODB.Recordset
            Dim PRD_CNT As Integer
            txtSalInv = Trim(RSINV(0))
            SQTY = RSINV(3)
            sinvdat.MinDate = "01/01/1999"
            sinvdat.MaxDate = "12/31/2015"
            sinvdat = RSINV(1)
            TTYP = RSINV(5)
            PID = RSINV(4)
            SID = RSINV(6)
            PrID = RSINV(2)
            
            If rsAvt.State = 1 Then rsAvt.Close
            Set rsAvt = Nothing
            'rsAvt.Open "SELECT MBL_IDY,PRD_IDY,CST_IDY,INV_IDY FROM AVT WHERE PRD_IDY IN(SELECT PRD_IDY FROM INV WHERE INV_IDY='" & Trim(txtSalInv) & "') and mbl_idy<>' ' and sal_typ is null or sal_typ='SRN' or sal_typ<>'SAL' order by mbl_idy", CON, adOpenKeyset, adlockpessimistic
            rsAvt.Open "SELECT MBL_IDY,PRD_IDY,CST_IDY,INV_IDY FROM AVT WHERE PRD_IDY IN(SELECT PRD_IDY FROM INV WHERE inv_dat=ctod('" & sinvdat & "') and INV_IDY='" & Trim(txtSalInv) & "') and mbl_idy<>' ' and sal_typ is null or sal_typ='SRN' or sal_typ<>'SAL' order by mbl_idy", con, adOpenKeyset, adLockPessimistic
            
            If pct_rst.State = 1 Then pct_rst.Close
            pct_rst.Open "seleCT prd.PRD_nme,SUM(inv.prd_qty),INV.PRD_IDY FROM INV,prd WHERE EXC_INF='Y' AND inv_dat=ctod('" & sinvdat & "') and prd.prd_idy=inv.prd_idy and TRIM(INV_IDY)='" & Trim(txtSalInv) & "' GROUP BY PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
            
            ppp = 0
            
            While pct_rst.EOF = False
            
                If pct_rst1.State = 1 Then pct_rst1.Close
                pct_rst1.Open "select count(prd_idy) from avt where sal_dat=ctod('" & sinvdat.Value & "') and sal_idy='" & Trim(txtSalInv) & "' and prd_idy='" & Trim(pct_rst(2)) & "'", con, adOpenKeyset, adLockPessimistic
                PRD_CNT = 0
                
                On Error GoTo sw
                If pct_rst1.EOF = False Then
                PRD_CNT = pct_rst(1) - pct_rst1(0)
                Else
                PRD_CNT = pct_rst(1)
                End If
                Ptno(ppp) = Trim(PRD_CNT) & " : " & Trim(pct_rst(0))
                Ptno(ppp).Visible = True
                pn.Caption = ppp
                pn.Visible = False
sw:
                pct_rst.MoveNext
                ppp = ppp + 1
            
            Wend
            
            pct_rst.Close
            Set pct_rst = Nothing
            
            lv.ColumnHeaders.CLEAR
            lv.ListItems.CLEAR
            lv.ColumnHeaders.Add , , "Mobile ID", 1800
            lv.ColumnHeaders.Add , , "SKU Name", 2500
            lv.ColumnHeaders.Add , , "Salesman", 0
            lv.ColumnHeaders.Add , , "Invoice ID", 0
            lv.ColumnHeaders.Add , , "ID", 0
            lv.ColumnHeaders.Add , , "Pr Id", 0
            lv.ColumnHeaders.Add , , "smn Id", 0
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_idy,prd_nme from prd", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act", con, adOpenKeyset, adLockPessimistic
            If RSSMN.State = 1 Then RSSMN.Close
            RSSMN.Open "select smn_idy,smn_nme from smn", con, adOpenKeyset, adLockPessimistic
            If RSSMN.RecordCount > 0 Then RSSMN.MoveFirst
            RSSMN.Find "smn_idy='" & SID & "'"
            If Not RSSMN.EOF Then txtSmn = RSSMN(1)
            If rsmar.State = 1 Then rsmar.Close
            rsmar.Open "SELECT mar.mar_idy,mar_nme,act_idy from mar,act where mar.mar_idy like act.mar_idy and act_idy='" & Trim(PID) & "'", con, adOpenKeyset, adLockPessimistic
            If Not rsmar.EOF Then txtMarket = rsmar(1)
            Dim SNO As Integer
            SNO = 1
            If rsAvt.RecordCount > 0 Then
                Do While Not rsAvt.EOF
                    If rsprd.State = 1 Then rsprd.Close
                    rsprd.Open "select prd_idy,prd_nme from PRD WHERE PRD_idy='" & Trim(rsAvt(1)) & "'", con, adOpenKeyset, adLockPessimistic
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_idy,act_nme from act WHERE act_idy='" & Trim(PID) & "'", con, adOpenKeyset, adLockPessimistic
                    Set lst = lv.ListItems.Add(, , rsAvt(0))    '1
                    If Not rsprd.EOF Then lst.ListSubItems.Add , , Trim(rsprd(1)) & ""
                    If RSSMN.RecordCount > 0 Then
'                    RSSMN.MoveFirst
'                    RSSMN.Find "smn_idy='" & sid & "'"
                    lst.ListSubItems.Add , , RSSMN(1)   '2
                    Else
                    lst.ListSubItems.Add , , " "
                    End If
                    lst.ListSubItems.Add , , rsAvt(3)
                    lst.ListSubItems.Add , , Trim(PID)
                    lst.ListSubItems.Add , , Trim(rsprd(0)) & ""
                    If Not rsprd.EOF Then
                    lst.ListSubItems.Add , , RSSMN(0) & ""
                    Else
                    lst.ListSubItems.Add , , " " & ""
                    End If
                    rsAvt.MoveNext
                    SNO = SNO + 1
                Loop
                txtPartyname = rsact(1)
                Label4 = SQTY 'RSINV(0)
                Label4.Visible = False
                If rsAvt.EOF = True Then
                    SSTab1.Tab = 1
                    lv.SetFocus
                    SSTab1.TabEnabled(3) = False
                End If
            Else
                MsgBox "Mobile Numbers not Found for the selected Invoice Product!", vbOKOnly, "DAS Version 4.0"
                SSTab1.Tab = 1
                txtSalInv = ""
                txtSmn = ""
                txtMarket = ""
                txtSMbl = ""
                txtSalInv.SetFocus
                SSTab1.TabEnabled(3) = False
            End If
Else

Dim rst_chk As New ADODB.Recordset

    If RSINV.State = 1 Then RSINV.Close
    RSINV.Open "select inv.inv_idy,inv.inv_dat,inv.prd_idy,SUM(inv.prd_qty),inl.cst_idy,inv.txn_typ,inl.smn_idy,exc_inf from inv,inl,prd where inv.inv_idy like inl.inv_idy and inv.inv_dat=inl.inv_dat AND inv.txn_typ in('SAL','STO','PRT') and inv.prd_idy like prd.prd_idy and prd.exc_inf='Y' and inv.stk_typ in('Salable','Free') group by inv.inv_idy,inv.inv_dat", con, adOpenKeyset, adLockPessimistic ' inv.inv_idy not in(select sal_idy from avt group by sal_idy) and
    
    If RSINV.RecordCount > 0 Then
        
        With LVP
        .ColumnHeaders.CLEAR
        .ListItems.CLEAR
        .Visible = True
        Label49.Caption = "Select An Invoice"
        .ColumnHeaders.Add , , "Invoice ID"
        .ColumnHeaders.Add , , "SKU Name"
        .ColumnHeaders.Add , , "Quantity"
        .ColumnHeaders.Add , , "Invoice Date"
        .ColumnHeaders.Add , , "Party", 1000
        .ColumnHeaders.Add , , "Type", 0
        .ColumnHeaders.Add , , "Salesman", 1000
        .ColumnHeaders.Add , , "SKU ID", 0
        .ColumnHeaders.Add , , "Party ID", 0
        .ColumnHeaders.Add , , "Salesman ID", 0
    End With
        Dim CHK_SAL As Boolean
        CHK_SAL = False
        
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_idy,prd_nme from prd", con, adOpenKeyset, adLockPessimistic
        While Not RSINV.EOF
            If rschk.State = 1 Then rschk.Close
            'Set rschk = Nothing
            rschk.Open "select sal_idy,sal_dat from avt order by sal_idy", con, adOpenKeyset, adLockPessimistic
            If rschk.RecordCount > 0 Then rschk.MoveFirst
            rschk.filter = "sal_idy='" & Trim(RSINV(0)) & "' and sal_dat=#" & RSINV(1) & "#"
            
            If rschk.RecordCount < RSINV(3) Then
                Set lst = LVP.ListItems.Add(, , RSINV(0))
                'rsprd.Find "prd_idy='" & Trim(RSINV(2)) & "'"
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_idy,prd_nme from prd where prd_idy='" & Trim(RSINV(2)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsprd.EOF Then lst.ListSubItems.Add , , rsprd(1)
                lst.ListSubItems.Add , , RSINV(3)
                lst.ListSubItems.Add , , RSINV(1)
                If rst_chk.State = 1 Then rst_chk.Close
                rst_chk.Open "select act_nme from act where act_idy='" & RSINV(4) & "'", con, adOpenKeyset, adLockPessimistic
                If rst_chk.EOF = False Then
                lst.ListSubItems.Add , , Trim(rst_chk(0))
                End If
                lst.ListSubItems.Add , , RSINV(5)
                If rst_chk.State = 1 Then rst_chk.Close
                rst_chk.Open "select smn_nme from smn where smn_idy='" & RSINV(6) & "'", con, adOpenKeyset, adLockPessimistic
                If rst_chk.EOF = False Then
                lst.ListSubItems.Add , , Trim(rst_chk(0))
                End If
                If Not rsprd.EOF Then lst.ListSubItems.Add , , Trim(rsprd(0)) & ""
                lst.ListSubItems.Add , , Trim(RSINV(4))
                lst.ListSubItems.Add , , Trim(RSINV(6))
                CHK_SAL = True
            End If
            RSINV.MoveNext
        Wend
        If CHK_SAL = True Then
            SSTab1.TabEnabled(3) = True
            SSTab1.Tab = 3
            LVP.SetFocus
        Else
       
        
        MsgBox "No Invoices Found.", vbOKOnly, "DAS Version 4.0"
        txtSalInv.SetFocus
        End If
    Else
        MsgBox "No Invoices Found.", vbOKOnly, "DAS Version 4.0"
        txtSalInv.SetFocus
    End If
End If
    End If
If rst_chk.State = 1 Then rst_chk.Close
    Set rst_chk = Nothing
SAL_ERR:
    
End Sub

Private Sub txtSMbl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim chk As Integer, CHK1 As Integer
CHK1 = 0
If KeyCode = 13 Then

    If Len(Trim(Val(txtSMbl.Text))) = 10 Then
       
        Dim ALL_CHK As Boolean
        Dim lv_chk As Boolean
        Dim int_no  As Integer
        ALL_CHK = False
        
        For i = 1 To lv.ListItems.count
            If Trim(lv.ListItems(i).Text) = Trim(txtSMbl.Text) Then
                int_no = i
                lv.ListItems(i).Checked = True
                lv_chk = list_chk(lv.ListItems(i).Index)
                ALL_CHK = True
            End If
        Next
        
        Dim temp As String
        Dim temp1 As String
        Dim temp2 As String
        Dim temp3 As String
        Dim temp4 As String
        Dim temp5 As String
        Dim temp6 As String
        Dim temp_chk As Integer
       
        If lv_chk = False Then
        If int_no <> 0 Then
        lv.ListItems(int_no).Checked = False
        MsgBox "Cannot Select More Than Invoiced for this SKU Starter Pack!", vbCritical, "DAS Version 4.0"
        lv.SetFocus
        End If
        txtSMbl.Text = ""
        Exit Sub
        End If
        
        For i = 1 To lv.ListItems.count
        If Trim(lv.ListItems(i).Text) = Trim(txtSMbl.Text) Then
        temp_chk = 0
        
        For i1 = 1 To lv.ListItems.count
        If lv.ListItems(i1).Checked = False Then
            temp_chk = i1
            Exit For
        End If
        Next
        
        If temp_chk = 0 Then
            Exit Sub
        End If
        
            If i < temp_chk Then
            Exit For
            Else
            temp = lv.ListItems(i)
            temp1 = lv.ListItems(i).SubItems(1)
            temp2 = lv.ListItems(i).SubItems(2)
            temp3 = lv.ListItems(i).SubItems(3)
            temp4 = lv.ListItems(i).SubItems(4)
            temp5 = lv.ListItems(i).SubItems(5)
    '        temp6 = LV.ListItems(I).SubItems(6)
        
            For K = i To temp_chk + 1 Step -1
            If K >= 2 Then
            lv.ListItems(K).Text = lv.ListItems(K - 1).Text
            lv.ListItems(K).SubItems(1) = lv.ListItems(K - 1).SubItems(1)
            lv.ListItems(K).SubItems(2) = lv.ListItems(K - 1).SubItems(2)
            lv.ListItems(K).SubItems(3) = lv.ListItems(K - 1).SubItems(3)
            lv.ListItems(K).SubItems(4) = lv.ListItems(K - 1).SubItems(4)
            lv.ListItems(K).SubItems(5) = lv.ListItems(K - 1).SubItems(5)
    '        lv.ListItems(K).SubItems(6) = lv.ListItems(K - 1).SubItems(6)
            End If
            Next
        
            If temp_chk > 0 Then
                lv.ListItems(temp_chk).Text = temp
                lv.ListItems(temp_chk).SubItems(1) = temp1
                lv.ListItems(temp_chk).SubItems(2) = temp2
                lv.ListItems(temp_chk).SubItems(3) = temp3
                lv.ListItems(temp_chk).SubItems(4) = temp4
                lv.ListItems(temp_chk).SubItems(5) = temp5
    '            lv.ListItems(temp_chk).SubItems(6) = temp6
                lv.ListItems(temp_chk).Checked = True
                lv.ListItems(i).Checked = False
            End If
                CHK1 = 1
            End If
            Exit For
        End If
        Next
        '---------------------------------
        
        
        If ALL_CHK = False Then
            MsgBox "Mobile Number Invalid! No Such Number", vbCritical, "DAS Version 4.0"
            txtSMbl.SetFocus
        Else
            lv.SetFocus
        End If
        
    Else
    
        temp_chk = 0
        
        For i1 = 1 To lv.ListItems.count
            If Mid(Trim(lv.ListItems(i1).Text), 1, Len(txtSMbl.Text)) = Trim(txtSMbl.Text) Then
                temp_chk = i1
                Exit For
            End If
        Next
        
        If temp_chk = 0 Then
            MsgBox "Mobile Number Invalid! No Such Number", vbCritical, "DAS Version 4.0"
            txtSMbl.SetFocus
            Exit Sub
        End If
        
        Dim SER As Integer
        
        SER = 0
        
        For i1 = 1 To lv.ListItems.count
        If lv.ListItems(i1).Checked = False Then
            SER = i1
            Exit For
        End If
        Next
        
        If SER <> 0 Then
            For i1 = 1 To lv.ListItems.count
            If Mid(lv.ListItems(i1).Text, 1, Len(txtSMbl.Text)) = Trim(txtSMbl.Text) Then
                temp = lv.ListItems(SER)
                temp1 = lv.ListItems(SER).SubItems(1)
                temp2 = lv.ListItems(SER).SubItems(2)
                temp3 = lv.ListItems(SER).SubItems(3)
                temp4 = lv.ListItems(SER).SubItems(4)
                temp5 = lv.ListItems(SER).SubItems(5)
                lv.ListItems(SER).Text = lv.ListItems(i1)
                lv.ListItems(SER).SubItems(1) = lv.ListItems(i1).SubItems(1)
                lv.ListItems(SER).SubItems(2) = lv.ListItems(i1).SubItems(2)
                lv.ListItems(SER).SubItems(3) = lv.ListItems(i1).SubItems(3)
                lv.ListItems(SER).SubItems(4) = lv.ListItems(i1).SubItems(4)
                lv.ListItems(SER).SubItems(5) = lv.ListItems(i1).SubItems(5)
                lv.ListItems(i1).Text = temp
                lv.ListItems(i1).SubItems(1) = temp1
                lv.ListItems(i1).SubItems(2) = temp2
                lv.ListItems(i1).SubItems(3) = temp3
                lv.ListItems(i1).SubItems(4) = temp4
                lv.ListItems(i1).SubItems(5) = temp5
                SER = SER + 1
            End If
            Next
        Else
            If SER = 0 Then SER = 1
            For i1 = 1 To lv.ListItems.count
            If Mid(lv.ListItems(i1).Text, 1, Len(txtSMbl.Text)) = Trim(txtSMbl.Text) Then
                temp = lv.ListItems(SER)
                temp1 = lv.ListItems(SER).SubItems(1)
                temp2 = lv.ListItems(SER).SubItems(2)
                temp3 = lv.ListItems(SER).SubItems(3)
                temp4 = lv.ListItems(SER).SubItems(4)
                temp5 = lv.ListItems(SER).SubItems(5)
                lv.ListItems(SER).Text = lv.ListItems(i1)
                lv.ListItems(SER).SubItems(1) = lv.ListItems(i1).SubItems(1)
                lv.ListItems(SER).SubItems(2) = lv.ListItems(i1).SubItems(2)
                lv.ListItems(SER).SubItems(3) = lv.ListItems(i1).SubItems(3)
                lv.ListItems(SER).SubItems(4) = lv.ListItems(i1).SubItems(4)
                lv.ListItems(SER).SubItems(5) = lv.ListItems(i1).SubItems(5)
                lv.ListItems(i1).Text = temp
                lv.ListItems(i1).SubItems(1) = temp1
                lv.ListItems(i1).SubItems(2) = temp2
                lv.ListItems(i1).SubItems(3) = temp3
                lv.ListItems(i1).SubItems(4) = temp4
                lv.ListItems(i1).SubItems(5) = temp5
                SER = SER + 1
            End If
            Next
        End If
    
    End If
End If

End Sub

Private Sub txtSMbl_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 42 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtStart_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(Val(txtStart))) = 10 Then
        If rschk.State = 1 Then rschk.Close
        rschk.Open "SELECT * FROM AVT WHERE (CHR_STS<>'Yes' or chr_sts='No' or chr_sts is null)", con, adOpenKeyset, adLockPessimistic
        If Not rschk.EOF Then
            If rschk.State = 1 Then rschk.Close
            rschk.Open "SELECT * FROM AVT WHERE MBL_IDY='" & Trim(txtStart) & "' and (CHR_STS<>'Yes' or chr_sts='No' or chr_sts is null)", con, adOpenKeyset, adLockPessimistic
            If rschk.EOF = True Then
                cmdGen.Enabled = True
                cmdGen.SetFocus
            Else
                MsgBox "Mobile Number Exists in Stock!", vbOKOnly, "DAS Version 4.0"
                txtStart.SetFocus
                SendKeys "{home}+{end}"
            End If
        Else
            cmdGen.Enabled = True
            cmdGen.SetFocus
        End If
    Else
        MsgBox "10 Digit Mobile Number Required!", vbOKOnly, "DAS Version 4.0"
        txtStart.SetFocus
        SendKeys "{home}+{end}"
    End If
End If



If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 42 Then
Else
    KeyAscii = 0
End If
If KeyAscii <> 27 Then
If Not Len(Trim(txtVch)) = 10 Then
MsgBox "Voucher ID Selection Required!", vbCritical, "DAS Version 4.0"
KeyAscii = 0
txtVch.SetFocus
End If
End If
End Sub

Private Sub txtStart_KeyPress(KeyAscii As Integer)
'If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 42 Then
'Else
'    KeyAscii = 0
'End If
'If KeyAscii <> 27 Then
'If Not Len(Trim(txtVch)) = 10 Then
'MsgBox "Voucher ID Selection Required!", vbCritical, "DAS Version 4.0"
'KeyAscii = 0
'txtVch.SetFocus
'End If
'End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
ElseIf Button.KEY = "s" Then
    If SSTab1.Tab = 0 Then
        If MS.Rows > 1 Then
            Command3_Click
        End If
    ElseIf SSTab1.Tab = 1 Then
        If lv.ListItems.count > 0 Then
            cmdsUpdate_Click
        End If
    ElseIf SSTab1.Tab = 2 Then
        If MSA.Rows > 1 Then
            CMDAUPDATE_Click
        End If
    ElseIf SSTab1.Tab = 5 Then
        If MSC.Rows > 1 Then
            cmdCUpdate_Click
        End If
    ElseIf SSTab1.Tab = 6 Then
        If MSR.Rows > 1 Then
            cmdRUpdate_Click
        End If
    End If
  
ElseIf Button.KEY = "c" Then
    If SSTab1.Tab = 0 Then
        inward.Visible = False
        MS.Rows = 1
        txtVch = ""
        txtStart = ""
        invdat = Date$
        txtVch.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        lv.ListItems.CLEAR
        txtPartyname = ""
        sinvdat = Date$
        txtSalInv = ""
        txtSalInv.SetFocus
    ElseIf SSTab1.Tab = 2 Then
        MSA.Rows = 1
        txtMobile = ""
        txtInvoice = ""
        txtRetailer = ""
        txtSalesman = ""
        'actDate.SetFocus
        'Option1(2).Value = True
        If chkAAll.Enabled = False Then chkAAll.Enabled = True
        chkAAll.SetFocus
        chkAAll.Value = 1
    ElseIf SSTab1.Tab = 5 Then
        MSC.Rows = 1
        txtcMobile = ""
        txtcInvoice = ""
        txtcRetailer = ""
        txtcSalesman = ""
        'cafdate.SetFocus
        'optDate.Value = True
        If chkCAll.Enabled = False Then chkCAll.Enabled = True
        chkCAll.SetFocus
        chkCAll.Value = 1
    ElseIf SSTab1.Tab = 6 Then
        MSR.Rows = 1
        txtrMobile = ""
        txtrInvoice = ""
        txtrRetailer = ""
        txtrSalesman = ""
        txtrMobile = ""
        If chkRAll.Enabled = False Then chkRAll.Enabled = True
        chkRAll.Value = 1
        chkRAll.SetFocus
    End If
ElseIf Button.KEY = "e" Then
    exitfrm
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If SSTab1.Tab = 2 Then
        If M_FRM_MBL.ActiveControl.name = "txtInvoice" Then
            txtInvoice_KeyDown 13, 0
        ElseIf M_FRM_MBL.ActiveControl.name = "txtRetailer" Then
            txtRetailer_KeyDown 13, 0
        ElseIf M_FRM_MBL.ActiveControl.name = "txtSalesman" Then
            txtSalesman_KeyDown 13, 0
        End If
    ElseIf SSTab1.Tab = 5 Then
        If M_FRM_MBL.ActiveControl.name = "txtcInvoice" Then
            txtcInvoice_KeyDown 13, 0
        ElseIf M_FRM_MBL.ActiveControl.name = "txtcRetailer" Then
            txtcRetailer_KeyDown 13, 0
        ElseIf M_FRM_MBL.ActiveControl.name = "txtcSalesman" Then
            txtcSalesman_KeyDown 13, 0
        End If
    ElseIf SSTab1.Tab = 6 Then
        If M_FRM_MBL.ActiveControl.name = "txtrInvoice" Then
            txtrInvoice_KeyDown 13, 0
        ElseIf M_FRM_MBL.ActiveControl.name = "txtrRetailer" Then
            txtrRetailer_KeyDown 13, 0
        ElseIf M_FRM_MBL.ActiveControl.name = "txtrSalesman" Then
            txtrSalesman_KeyDown 13, 0
        End If
    ElseIf SSTab1.Tab = 0 Then
        Label49 = "Select Invoice"
        invlist
    ElseIf SSTab1.Tab = 1 Then
        sale_view
    End If
ElseIf Button.KEY = "d" Then
ElseIf Button.KEY = "pre" Then
    
ElseIf Button.KEY = "p" Then
ElseIf Button.KEY = "pr" Then
ElseIf Button.KEY = "crt" Then
    
End If
End Sub

Private Sub txttemp_KeyDown(KeyCode As Integer, Shift As Integer)
Dim chk_inn As Boolean
Dim chk, RON As Integer
RON = MS.Row
If rsAvt.State = 1 Then rsAvt.Close
rsAvt.Open "select mbl_idy from avt where mbl_idy = '" & Trim(txtTemp) & "' AND TRIM(Chr_sts) IS NULL", con, adOpenKeyset
'MsgBox "" & rsAvt("chr_STS")
If KeyCode = 13 Then
    If Len(Trim(txtTemp)) = 10 Then
      'If Not rsAvt.EOF Then rsAvt.Find "mbl_idy='" & Trim(txtTemp) & "'"
        If Not rsAvt.RecordCount > 0 Then
            chk = 0
            For r = 1 To MS.Rows - 1 'ROWNO - 1 'MS.Rows - 1
                If Trim(MS.TextMatrix(r, 3)) = Trim(txtTemp) Then
                    MsgBox "Mobile Number Already Exists!", vbOKOnly, "DAS Version 4.0"
                    chk = 1
                    txtTemp.SetFocus
                    SendKeys "{home}+{end}"
                    Exit For
                End If
            Next
            If chk = 0 Then
                MS.Text = Trim(txtTemp)
                txtTemp.Visible = False
                t2 = 1
                Dim ppp As Integer
                ppp = MS.Row + 1
                MS.SetFocus
'                While ppp <= MS.Rows - 1
'
'                    chk_inn = False
'
'                    For r = 1 To ROWNO - 1
'                        If Val(Val(txtTemp) + t2) = Val(MS.TextMatrix(r, 3)) Then
'                            chk_inn = True
'                        End If
'                    Next
'                    If chk_inn = False Then
'                        MS.Row = ppp
'                        MS.Col = 3
'                        MS.Text = Val(txtTemp) + t2
'                        ppp = ppp + 1
'                    End If
'                    t2 = t2 + 1
'                Wend
'                If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
''                MS.Row = MS(R+1)
'                MS.Col = 3
'                MS.SetFocus
      '       End If
            End If
        Else
            MsgBox "Mobile Number Already Exists!", vbOKOnly, "DAS Version 4.0"
            txtTemp.SetFocus
            SendKeys "{home}+{end}"
        End If
    Else
        MsgBox "10 Digit Mobile Number Required!", vbOKOnly, "DAS Version 4.0"
        txtTemp.SetFocus
        SendKeys "{home}+{end}"
    End If
    
End If
End Sub

Private Sub txtTemp_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 42 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtTemp_LostFocus()
'            For r = 1 To MS.Rows - 1
'                If Trim(MS.TextMatrix(r, 3)) = Trim(txtTemp) Then
'                    MsgBox "Mobile Number Already Exists!", vbOKOnly, "DAS Version 4.0"
'                    chk = 1
'                    txtTemp.SetFocus
'                    SendKeys "{home}+{end}"
'                    Exit For
'                End If
'            Next

End Sub

Public Sub Txtvch_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("inv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
If K = 13 Or K = 40 Then
    Label49 = "Select Invoice"
    invlist
End If
End Sub
Private Sub updateavt()
' Updating records with Mobile Numbers
If MsgBox("Save Mobile Numbers ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If

If Mid(Trim(txtVch), 1, 3) = "PUR" Or Mid(Trim(txtVch), 1, 3) = "OPS" Or Mid(Trim(txtVch), 1, 3) = "STI" Or Mid(Trim(txtVch), 1, 3) = "SRN" Then
    If MS.Rows <= 1 Then
        MsgBox "Select Opening Stock or Purchase Transactions to Update!", vbInformation + vbOKOnly, "DAS Version 4.0"
        Exit Sub
    End If
    For r = 1 To MS.Rows - 1
        If Val(MS.TextMatrix(r, 3)) <= 0 Or Len(Trim(Val(MS.TextMatrix(r, 3)))) < 10 Then
            MsgBox "10 Digit Mobile Number Required!", vbCritical + vbOKOnly, "DAS Version 4.0"
            If txtTemp.Visible = True Then txtTemp.Visible = False
            MS.Row = r
            MS.Col = 3
            MS.SetFocus
            Exit Sub
        End If
    Next
    If rschk.State = 1 Then rschk.Close
    Set rschk = Nothing
    rschk.Open "select mbl_idy,txn_typ from avt where txn_typ in ('PUR','OPS','SRN','STI') and (chr_sts<>'Yes' or chr_sts='No' or chr_sts is null)", con, adOpenKeyset, adLockPessimistic
    If Not rschk.EOF Then
            For r = 1 To MS.Rows - 1
                If rschk.RecordCount > 0 Then rschk.MoveFirst
                rschk.Find "MBL_IDY='" & Trim(MS.TextMatrix(r, 3)) & "'"
                If Not rschk.EOF Then
                    MsgBox "Mobile Number Already Exists", vbInformation + vbOKOnly, "DAS Version 4.0"
                    MS.Row = r
                    MS.Col = 3
                    MS.SetFocus
                    Exit Sub
                End If
            Next
            If rsdir.State = 1 Then rsdir.Close
            Set rsdir = Nothing
            rsdir.Open "select dir_als,dir_msg from dir where dir_idy='CUR' and dir_idx='FYR'", con, adOpenKeyset, adLockPessimistic
            con.Execute "delete from avt where inv_idy='" & Trim(txtVch) & "' and inv_dat=ctod('" & invdat.Value & "') and (mbl_idy is null or mbl_idy='')"
            'CON.Execute "delete from avt where inv_idy='" & Trim(txtVch) & "' and inv_dat>=ctod('" & Format(CDate(rsdir("dir_als")), "mm/DD/yyyy") & "') and inv_dat<=ctod('" & Format(CDate(rsdir("dir_msg")), "mm/DD/yyyy") & "') and (mbl_idy is null or mbl_idy='')"
        If rsAvt.State = 1 Then rsAvt.Close
        Set rsAvt = Nothing
        rsAvt.Open "SELECT * FROM AVT", con, adOpenKeyset, adLockPessimistic
        For r = 1 To MS.Rows - 1
            
            rsAvt.AddNew
            rsAvt("INV_IDY") = Trim(txtVch) & ""
            rsAvt("INV_DAT") = invdat.Value
            rsAvt("PRD_IDY") = MS.TextMatrix(r, 5)
            rsAvt("CAT_IDY") = MS.TextMatrix(r, 8) & ""
            rsAvt("MBL_IDY") = Trim(MS.TextMatrix(r, 3)) & ""
            rsAvt("SPL_IDY") = MS.TextMatrix(r, 7) & ""
            rsAvt("TXN_TYP") = MS.TextMatrix(r, 6)
            rsAvt.Update
        
        Next
    End If
Else
    If LVSRN.ListItems.count = 0 Then
        MsgBox "No Invoice Found", vbInformation + vbOKOnly, "DAS Version 4.0"
        txtVch.SetFocus
    Exit Sub
    End If
    Dim cn As Integer
    If RSINV.State = 1 Then RSINV.Close
    Set RSINV = Nothing
    RSINV.Open "SELECT iNV.PRD_QTY,inv.prd_idy,prd.prd_nme,exc_inf FROM INV,prd WHERE INV_IDY='" & Trim(txtVch) & "' and inv.prd_idy=prd.prd_idy and exc_inf='Y'", con, adOpenStatic
    While Not RSINV.EOF
        cn = 0
        For r = 1 To LVSRN.ListItems.count
            If LVSRN.ListItems(r).Checked = True Then
                cn = cn + 1
            End If
        Next
        If RSINV(0) <> cn Then
            MsgBox "Cannot Save! " & Trim(RSINV(2)) & " Salereturn Quantity is: " & RSINV(0), vbInformation, "DAS Version 4.0"
            LVSRN.SetFocus
            Exit Sub
        End If
        RSINV.MoveNext
    Wend
    If rsAvt.State = 1 Then rsAvt.Close
    Set rsAvt = Nothing
    rsAvt.Open "select * from avt", con, adOpenKeyset, adLockPessimistic
    con.Execute "delete from avt where sal_typ='SRN'"
    For r = 1 To LVSRN.ListItems.count
        If LVSRN.ListItems(r).Checked = True Then
            If rsAvt.RecordCount > 0 Then rsAvt.MoveFirst
            rsAvt.filter = "mbl_idy='" & Trim(LVSRN.ListItems(r).Text) & "'"
            If Not rsAvt.EOF Then
                rsAvt("sal_typ") = "SRN"
                rsAvt.Update
            End If
        End If
    Next
End If
    cancel
End Sub
Private Sub exitfrm()
Unload Me
End Sub
Private Sub cancel()
inward.Visible = False
For s = 0 To 5
    If Ptno(s).Visible = True Then Ptno(s).Visible = False
Next
For Each Control In Me.Controls
    If TypeOf Control Is TextBox Then
    Control.Text = ""
    Control.Enabled = True
    End If
Next
txtRetName = ""
invdat = Date$
MS.Rows = 1
If lv.ListItems.count > 0 Then lv.ListItems.CLEAR
If MSA.Rows > 1 Then MSA.Rows = 1
If MSC.Rows > 1 Then MSC.Rows = 1
If MSR.Rows > 1 Then MSR.Rows = 1
If txtTemp.Visible = True Then txtTemp.Visible = False
If LVSRN.Visible = True Then LVSRN.Visible = False
MS.Visible = True
SSTab1.Tab = 0
txtVch.SetFocus
SSTab1.TabEnabled(3) = False
cmbType.ListIndex = 0
txtRet.Enabled = True
txtRet = ""
txtRetName = ""
If LVF.ListItems.count > 0 Then LVF.ListItems.CLEAR
If LVR.ListItems.count > 0 Then LVR.ListItems.CLEAR
'Retailer returns
LVMOD.ListItems.CLEAR
txtModRet.Text = ""
txtModRetName.Caption = ""
txtModInv.Text = ""
'Retailer returns
'Churn
optrSmn.Value = False
optrRet.Value = False
optrInvoice.Value = False
optrMobile.Value = False
'churn
txtcInvoice.Visible = True
txtcSalesman.Visible = True
txtcRetailer.Visible = True
txtcMobile.Visible = True

optSmn.Enabled = True
optInvoice.Enabled = True
optRet.Enabled = True
optMobile.Enabled = True
'Caf
optSmn.Value = False
optInvoice.Value = False
optRet.Value = False
optMobile.Value = False
txtrSalesman.Visible = True
txtrRetailer.Visible = True
txtrInvoice.Visible = True
txtrMobile.Visible = True
'Caf

'enable
M_FRM_OPT.Enabled = True
Option1(0).Enabled = True
Option1(1).Enabled = True
Option1(3).Enabled = True
Option1(4).Enabled = True
txtSalesman.Enabled = True
txtRetailer.Enabled = True
txtInvoice.Enabled = True
txtMobile.Enabled = True
'enable
    lblno(0).Visible = False
    lblno(3).Visible = False
    lblno(2).Visible = False
    lblno(1).Visible = False
'enable
optrSmn.Enabled = True
optrRet.Enabled = True
optrInvoice.Enabled = True
optrMobile.Enabled = True
txtrSalesman.Enabled = True
txtSalesman.Visible = True
txtRetailer.Visible = True
txtInvoice.Visible = True
txtMobile.Visible = True
txtrRetailer.Enabled = True
txtrInvoice.Enabled = True
txtrMobile.Enabled = True
'---------------
lblno(0).Visible = False
lblno(3).Visible = False
lblno(2).Visible = False
lblno(1).Visible = False
'---------------
If MSC.Rows > 1 Then MSC.Rows = 1
If MSR.Rows > 1 Then MSR.Rows = 1
Frame1.Enabled = True
End Sub
Private Sub invlist()
If rsAvt.State = 1 Then rsAvt.Close
Set rsAvt = Nothing
    rsAvt.Open "select inv_idy,inv_dat,prd_idy,spl_idy from avt where (mbl_idy is null or mbl_idy='') group by inv_idy,inv_dat", con, adOpenStatic
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "select prd_idy,prd_nme from prd", con, adOpenKeyset, adLockPessimistic
    CSTID = ""
    If rsAvt.RecordCount > 0 Then CSTID = Trim(rsAvt(3))
    If rsact.State = 1 Then rsact.Close
    Set rsact = Nothing
    rsact.Open "select act_idy,act_nme from act where act_idy like '" & Trim(CSTID) & "'", con, adOpenKeyset, adLockPessimistic
    If rsAvt.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
        LVP.ListItems.CLEAR
        LVP.ColumnHeaders.Add , , "Invoice ID", 1500
        LVP.ColumnHeaders.Add , , "Invoice Date", 1500
        LVP.ColumnHeaders.Add , , "SKU Name", 2500
        LVP.ColumnHeaders.Add , , "Supplier Name", 2000
   
        While Not rsAvt.EOF
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_idy,prd_nme from prd where prd_idy='" & Trim(rsAvt(2)) & "'", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act where act_idy='" & Trim(rsAvt(3)) & "' and act_idy like '" & Trim(CSTID) & "'", con, adOpenKeyset, adLockPessimistic
            Set lst = LVP.ListItems.Add(, , rsAvt(0))
            lst.ListSubItems.Add , , rsAvt(1)
            If Not rsprd.EOF Then lst.ListSubItems.Add , , rsprd(1)
            If Not rsact.EOF Then lst.ListSubItems.Add , , rsact(1)
            rsAvt.MoveNext
           
        Wend
        SSTab1.TabEnabled(3) = True
        SSTab1.Tab = 3
        LVP.Visible = True
        LVP.SetFocus
    End If
End Sub

Private Sub invlist_RETURNS()
    Dim RST_AVT As New ADODB.Recordset
    
    If rsAvt.State = 1 Then rsAvt.Close
    Set rsAvt = Nothing
    
    '=========================================================
        
        If chkModDate.Value = 1 Then DQRY = " and INL.INV_DAT=CTOD('" & modDate & "')"
'        If rsMod.State = 1 Then rsMod.Close
'        rsMod.Open "select inl.inv_idy,inl.cst_idy,inl.inv_dat from inl,inv,prd where inl.inv_idy like inv.inv_idy and inv.prd_idy like prd.prd_idy and prd.exc_inf in('y','Y') and cst_idy like '" & Trim(txtModRet) & "%'" & dqry & " group by inl.inv_idy", CON, adOpenStatic

    '=========================================================
    
    If RST_AVT.State = 1 Then RST_AVT.Close
    If Len(Trim(DQRY)) <> 0 Then
        RST_AVT.Open "select inv_idy,inv_dat from INL WHERE cst_idy like '" & Trim(txtModRet) & "%'" & DQRY & " group by inv_idy,inv_dat", con, adOpenStatic
    Else
        RST_AVT.Open "select inv_idy,inv_dat from INL WHERE cst_idy like '" & Trim(txtModRet) & "%'group by inv_idy,inv_dat", con, adOpenStatic
    End If
    If RST_AVT.EOF = False Then
        If rsAvt.State = 1 Then rsAvt.Close
        rsAvt.Open "select inv_idy,inv_dat,prd_idy,spl_idy from AVT WHERE INV_IDY='" & RST_AVT(0) & "' and inv_dat=CTOD('" & RST_AVT(1) & "') group by inv_idy,inv_dat", con, adOpenStatic
    Else
        MsgBox "Invoices not found!", vbInformation, "DAS Version 4.0"
        txtModInv.SetFocus
        Exit Sub
    End If
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "select prd_idy,prd_nme from prd", con, adOpenKeyset, adLockPessimistic
    CSTID = ""
    If rsAvt.RecordCount > 0 Then CSTID = Trim(rsAvt(3))
    If rsact.State = 1 Then rsact.Close
    Set rsact = Nothing
    rsact.Open "select act_idy,act_nme from act where act_idy like '" & Trim(CSTID) & "'", con, adOpenKeyset, adLockPessimistic
    If rsAvt.RecordCount > 0 Then
        LVP.ColumnHeaders.CLEAR
        LVP.ListItems.CLEAR
        LVP.ColumnHeaders.Add , , "Invoice ID", 1500
        LVP.ColumnHeaders.Add , , "Invoice Date", 1500
        LVP.ColumnHeaders.Add , , "SKU Name", 2500
        LVP.ColumnHeaders.Add , , "Supplier Name", 2000
   
        While Not rsAvt.EOF
            If rsprd.State = 1 Then rsprd.Close
            rsprd.Open "select prd_idy,prd_nme from prd where prd_idy='" & Trim(rsAvt(2)) & "'", con, adOpenKeyset, adLockPessimistic
            If rsact.State = 1 Then rsact.Close
            rsact.Open "select act_idy,act_nme from act where act_idy='" & Trim(rsAvt(3)) & "' and act_idy like '" & Trim(CSTID) & "'", con, adOpenKeyset, adLockPessimistic
            Set lst = LVP.ListItems.Add(, , rsAvt(0))
            lst.ListSubItems.Add , , rsAvt(1)
            If Not rsprd.EOF Then lst.ListSubItems.Add , , rsprd(1)
            If Not rsact.EOF Then lst.ListSubItems.Add , , rsact(1)
            rsAvt.MoveNext
           
        Wend
        SSTab1.TabEnabled(3) = True
        SSTab1.Tab = 3
        LVP.Visible = True
        LVP.SetFocus
    End If
End Sub


Private Sub Ynlist_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSA.Text = Ynlist.Text
    Ynlist.Visible = False
    If MSA.Row <> MSA.Rows - 1 Then MSA.Row = MSA.Row + 1
    MSA.SetFocus
End If
End Sub
Private Sub clrall()
MSA.Rows = 1
txtMobile = ""
txtInvoice = ""
actDate = Date$
txtRetailer = ""
txtSalesman = ""
Ynlist.Visible = False
'actDate.SetFocus
'Option1(2).Value = True
If chkAAll.Enabled = False Then chkAAll.Enabled = True
chkAAll.SetFocus
chkAAll.Value = 1
End Sub
Private Sub CLRCLL()
MSC.Rows = 1
txtcMobile = ""
txtcInvoice = ""
cafdate = Date$
txtcRetailer = ""
txtcSalesman = ""
CYnlist.Visible = False
If chkCAll.Enabled = False Then chkCAll.Enabled = True
chkCAll.SetFocus
chkCAll.Value = 1
End Sub
Private Sub CLRRLL()
MSR.Rows = 1
txtrMobile = ""
txtrInvoice = ""
ChrDate = Date$
txtrRetailer = ""
txtrSalesman = ""
RYnlist.Visible = False
If chkRAll.Enabled = False Then chkRAll.Enabled = True
chkRAll.SetFocus
chkRAll.Value = 1
End Sub
Private Sub TXTret_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
If rs.State = 1 Then rs.Close
        rs.Open " select * from act where act_idy like '" & Sincrement(txtRet) & "' and grp_clf='Debtor' and cst_typ in ('Retailer','Key A/c','Local Key A/c','Regnl Key A/c')", con, adOpenStatic
        If rs.RecordCount > 0 Then
            txtRet = Sincrement(txtRet)
            txtRetName = Trim(rs("ACT_nme"))
            dispDisMbl
            dispRetMbl (Trim(txtRet))
            'txtRet.Enabled = False
        Else
            actlist
        End If
End If
End Sub

Private Sub actlist()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
Dim ls As ListItem
rs.Open "select * from act where act_nme like '" & Trim(txtRet) & "%' and grp_clf='Debtor' and cst_typ in ('Retailer','Key A/c','Local Key A/c','Regnl Key A/c')", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    LVP.ColumnHeaders.CLEAR
    LVP.ColumnHeaders.Add , , "Party", 2500
    LVP.ColumnHeaders.Add , , "ID", 1500
    LVP.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = LVP.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy"))
        rs.MoveNext
    Next
    Label49 = "Select Retailer"
    SSTab1.Tab = 3
    LVP.Visible = True
    LVP.SetFocus
Else
    MsgBox "Retailer Not Found, Define Retailer !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 4
    txtRet.Enabled = True
    txtRet.SetFocus
End If
End Sub

Private Sub dispRetMbl(retid As String)
On Error GoTo errh
Dim rspr As New ADODB.Recordset
If rsEdt.State = 1 Then rsEdt.Close
rsEdt.Open "select mbl_idy,smn_idy,cst_idy,sal_dat,prd_idy from avt where sal_typ='SAL' and cst_idy='" & Trim(retid) & "' and (avt_sts is null or avt_sts='No' or avt_sts<>'Yes') order by mbl_idy", con, adOpenKeyset, adLockPessimistic
Dim ls As ListItem
LVF.ListItems.CLEAR
LVF.ColumnHeaders.CLEAR
LVF.ColumnHeaders.Add , , "Mobile ID", 2000
LVF.ColumnHeaders.Add , , "SKU Name", 3000
'LVF.ColumnHeaders.Add , , "Invoice Date", 1000
If rspr.State = 1 Then rspr.Close
rspr.Open "select prd_idy,prd_nme from prd", con, adOpenKeyset, adLockPessimistic
If rsEdt.RecordCount > 0 Then
    While Not rsEdt.EOF
        Set ls = LVF.ListItems.Add(, , Trim(rsEdt(0)))
        If rspr.State = 1 Then rspr.Close
        rspr.Open "select prd_idy,prd_nme from prd where prd_idy='" & Trim(rsEdt(4)) & "'", con, adOpenKeyset, adLockPessimistic
        If Not rspr.EOF Then
            ls.ListSubItems.Add , , Trim(rspr(1))
        Else
            ls.ListSubItems.Add , , ""
        End If
        rsEdt.MoveNext
    Wend
    LVF.SetFocus
End If
Exit Sub
errh:
MsgBox "Cannot Select Transaction!", vbInformation + vbOKOnly, "DAS Version 4.0"
End Sub
Private Sub LVF_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If LVF.SortOrder = lvwAscending Then
LVF.SortOrder = lvwDescending
ElseIf LVF.SortOrder = lvwDescending Then
LVF.SortOrder = lvwAscending
End If
LVF.SortKey = CH.Index - 1
End Sub

Private Sub LVF_DblClick()
'LVF_KeyDown 13, 0
End Sub
Private Sub LVR_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If LVR.SortOrder = lvwAscending Then
LVR.SortOrder = lvwDescending
ElseIf LVR.SortOrder = lvwDescending Then
LVR.SortOrder = lvwAscending
End If
LVR.SortKey = CH.Index - 1
End Sub

Private Sub LVR_DblClick()
'LVR_KeyDown 13, 0
End Sub

Private Sub dispDisMbl()
Dim rspr As New ADODB.Recordset
If rsEdt.State = 1 Then rsEdt.Close
rsEdt.Open "select mbl_idy,inv_dat,prd_idy from avt where (sal_typ IS NULL or sal_typ<>'SAL') and (mbl_idy is not null or mbl_idy=' ')order by mbl_idy", con, adOpenKeyset, adLockPessimistic
Dim ls As ListItem
LVR.ListItems.CLEAR
LVR.ColumnHeaders.CLEAR
LVR.ColumnHeaders.Add , , "Mobile ID", 2000
LVR.ColumnHeaders.Add , , "SKU Name", 3000
'LVR.ColumnHeaders.Add , , "Invoice Date", 1000

If rspr.State = 1 Then rspr.Close
rspr.Open "select prd_idy,prd_nme from prd", con, adOpenKeyset, adLockPessimistic

While Not rsEdt.EOF
    Set ls = LVR.ListItems.Add(, , Trim(rsEdt(0)) & "")
    If rspr.RecordCount > 0 Then rspr.MoveFirst
    rspr.filter = "prd_idy='" & Trim(rsEdt(2)) & "'"
    If Not rspr.EOF Then
        ls.ListSubItems.Add , , Trim(rspr(1)) & ""
    Else
        ls.ListSubItems.Add , , ""
    End If
    'ls.ListSubItems.Add , , rsEdt(1)
    rsEdt.MoveNext
Wend
End Sub
Private Sub LOADMSPC()
With MSPC
    .Rows = 1
    .Cols = 4
    .TextMatrix(0, 0) = "Invoice No"
    .ColWidth(0) = 1400
    .TextMatrix(0, 1) = "Supplier"
    .ColWidth(1) = 2500
    .TextMatrix(0, 2) = "Invoice Date"
    .ColWidth(2) = 1200
    .TextMatrix(0, 3) = "Mobile ID"
    .ColWidth(3) = 2500
    .ColAlignment(3) = 7
End With
End Sub
Private Sub dispPurInv()
Dim r As Integer
If rsEdt.State = 1 Then rsEdt.Close
rsEdt.Open "select INV_IDY,spl_idy,inv_dat,mbl_idy,act_nme from avt,act where (sal_typ is null or sal_typ<>'SAL') and mbl_idy is not null and avt.spl_idy like act.act_idy", con, adOpenKeyset, adLockPessimistic
MSPC.Rows = MSPC.Rows + rsEdt.RecordCount
r = 1
While Not rsEdt.EOF
    MSPC.TextMatrix(r, 0) = Trim(rsEdt(0)) & ""
    MSPC.TextMatrix(r, 1) = Trim(rsEdt(4)) & ""
    MSPC.TextMatrix(r, 2) = Format(Trim(rsEdt(2)) & "", "dd/MM/yy")
    MSPC.TextMatrix(r, 3) = Trim(rsEdt(3)) & ""
    r = r + 1
    rsEdt.MoveNext
Wend
If rsEdt.State = 1 Then rsEdt.Close
rsEdt.Open "select INV_IDY,spl_idy,inv_dat,mbl_idy from avt where (sal_typ is null or sal_typ<>'SAL') and mbl_idy is not null and txn_typ='OPS'", con, adOpenKeyset, adLockPessimistic
MSPC.Rows = MSPC.Rows + rsEdt.RecordCount
While Not rsEdt.EOF
    MSPC.TextMatrix(r, 0) = Trim(rsEdt(0)) & ""
    MSPC.TextMatrix(r, 1) = ""
    MSPC.TextMatrix(r, 2) = Format(Trim(rsEdt(2)) & "", "dd/MM/yy")
    MSPC.TextMatrix(r, 3) = Trim(rsEdt(3)) & ""
    r = r + 1
    rsEdt.MoveNext
Wend
End Sub
Private Sub actlistMod()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
Dim ls As ListItem
rs.Open "select * from act where grp_clf='Debtor' and cst_typ in ('Retailer','Key A/c','Local Key A/c','Regnl Key A/c')", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    LVP.ColumnHeaders.CLEAR
    LVP.ColumnHeaders.Add , , "Party", 2500
    LVP.ColumnHeaders.Add , , "ID", 1500
    LVP.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = LVP.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy"))
        rs.MoveNext
    Next
    Label49 = "Select Party Account (Retail Circulation)"
    SSTab1.Tab = 3
    LVP.Visible = True
    LVP.SetFocus
Else
    MsgBox "Retailer Not Found, Define Retailer !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 7
    txtModRet.Enabled = True
    txtModRet.SetFocus
End If
End Sub


Private Sub actlistMod1()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
Dim ls As ListItem
rs.Open "select * from act where grp_clf='Creditor' and cst_typ in ('Retailer','Key A/c','Local Key A/c','Regnl Key A/c') and act_typ='PUR'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    LVP.ColumnHeaders.CLEAR
    LVP.ColumnHeaders.Add , , "Party", 2500
    LVP.ColumnHeaders.Add , , "ID", 1500
    LVP.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = LVP.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy"))
        rs.MoveNext
    Next
    Label49 = "Select Party Account (Retail Circulation)"
    SSTab1.Tab = 3
    LVP.Visible = True
    LVP.SetFocus
Else
    MsgBox "Retailer Not Found, Define Retailer !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 7
    txtModRet.Enabled = True
    txtModRet.SetFocus
End If
End Sub


Public Sub DF6()
If M_FRM_MBL.ActiveControl.name = "LVF" Then
    For r = 1 To LVF.ListItems.count
        If LVF.ListItems(r).Checked = True Then
            LVR.SetFocus
            Exit Sub
        End If
    Next
    MsgBox "Select Numbers to be Replaced!", vbInformation + vbOKOnly, "DAS Version 4.0"
    LVF.SetFocus
ElseIf M_FRM_MBL.ActiveControl.name = "LVR" Then
    LVF.SetFocus
    Exit Sub
End If
End Sub

Public Sub sale_view()
' If RSINV.State = 1 Then RSINV.Close
' RSINV.Open "select inv.inv_idy,inv.inv_dat,inv.prd_idy,SUM(inv.prd_qty),inl.cst_idy,inv.txn_typ,inl.smn_idy,exc_inf from inv,inl,prd where inv.inv_idy like inl.inv_idy and inv.inv_dat=inl.inv_dat AND inv.txn_typ in('SAL','STO','PRT') and inv.prd_idy like prd.prd_idy and TRIM(INV.STK_UPD)= 'Y' AND prd.exc_inf='Y' and inv.stk_typ in('Salable','Free') group by inv.inv_idy", CON, adOpenKeyset, adlockpessimistic ' inv.inv_idy not in(select sal_idy from avt group by sal_idy) and
'    With LVP
'        .ColumnHeaders.CLEAR
'        .ListItems.CLEAR
'        .Visible = True
'        Label49.Caption = "Select An Invoice"
'        .ColumnHeaders.Add , , "Invoice ID"
'        .ColumnHeaders.Add , , "SKU Name"
'        .ColumnHeaders.Add , , "Quantity"
'        .ColumnHeaders.Add , , "Invoice Date"
'        .ColumnHeaders.Add , , "Party", 0
'        .ColumnHeaders.Add , , "Type", 0
'        .ColumnHeaders.Add , , "Salesman", 0
'        .ColumnHeaders.Add , , "SKU ID", 0
'    End With
'    If RSINV.RecordCount > 0 Then
'        If rsprd.State = 1 Then rsprd.Close
'        rsprd.Open "select prd_idy,prd_nme from prd", CON, adOpenKeyset, adlockpessimistic
'
'        While Not RSINV.EOF
'            If rschk.State = 1 Then rschk.Close
'    rschk.Open "select sal_idy,sal_dat from avt where sal_idy='" & Trim(RSINV(0)) & " ' and sal_dat='" & RSINV(1) & "' order by sal_idy", CON, adOpenKeyset, adlockpessimistic
'            If rschk.EOF Then
''                If rsprd.RecordCount > 0 Then rsprd.MoveFirst
'                Set lst = LVP.ListItems.Add(, , RSINV(0))
'                If rsprd.State = 1 Then rsprd.Close
'                rsprd.Open "select prd_idy,prd_nme from prd where prd_idy='" & Trim(RSINV(2)) & "'", CON, adOpenKeyset, adlockpessimistic
'                'rsprd.Find "prd_idy='" & Trim(RSINV(2)) & "'"
'                If Not rsprd.EOF Then lst.ListSubItems.Add , , rsprd(1)
'                lst.ListSubItems.Add , , RSINV(3)
'                lst.ListSubItems.Add , , RSINV(1)
'                lst.ListSubItems.Add , , RSINV(4)
'                lst.ListSubItems.Add , , RSINV(5)
'                lst.ListSubItems.Add , , Trim(RSINV(6)) & ""
'                If Not rsprd.EOF Then lst.ListSubItems.Add , , Trim(rsprd(0)) & ""
'            End If
'            RSINV.MoveNext
'        Wend
'        SSTab1.TabEnabled(3) = True
'        SSTab1.Tab = 3
'        LVP.SetFocus
'    Else
'        MsgBox "No Invoices Found.", vbOKOnly, "DAS Version 4.0"
'        txtSalInv.SetFocus
'    End If
' ABOVE CODE COMMENTED BY PKB AND PLACED THE BELOW CODES
Dim rst_chk As New ADODB.Recordset

    If RSINV.State = 1 Then RSINV.Close
    RSINV.Open "select inv.inv_idy,inv.inv_dat,inv.prd_idy,SUM(inv.prd_qty),inl.cst_idy,inv.txn_typ,inl.smn_idy,exc_inf from inv,inl,prd where inv.inv_idy like inl.inv_idy and inv.inv_dat=inl.inv_dat AND inv.txn_typ in('SAL','STO','PRT') and inv.prd_idy like prd.prd_idy and prd.exc_inf='Y' and inv.stk_typ in('Salable','Free') group by inv.inv_idy,inv.inv_dat", con, adOpenKeyset, adLockPessimistic ' inv.inv_idy not in(select sal_idy from avt group by sal_idy) and
    
    If RSINV.RecordCount > 0 Then
        
        With LVP
        .ColumnHeaders.CLEAR
        .ListItems.CLEAR
        .Visible = True
        Label49.Caption = "Select An Invoice"
        .ColumnHeaders.Add , , "Invoice ID"
        .ColumnHeaders.Add , , "SKU Name"
        .ColumnHeaders.Add , , "Quantity"
        .ColumnHeaders.Add , , "Invoice Date"
        .ColumnHeaders.Add , , "Party", 1000
        .ColumnHeaders.Add , , "Type", 0
        .ColumnHeaders.Add , , "Salesman", 1000
        .ColumnHeaders.Add , , "SKU ID", 0
        .ColumnHeaders.Add , , "Party ID", 0
        .ColumnHeaders.Add , , "Salesman ID", 0
    End With
        Dim CHK_SAL As Boolean
        CHK_SAL = False
        
        If rsprd.State = 1 Then rsprd.Close
        rsprd.Open "select prd_idy,prd_nme from prd", con, adOpenKeyset, adLockPessimistic
        While Not RSINV.EOF
            If rschk.State = 1 Then rschk.Close
            'Set rschk = Nothing
            rschk.Open "select sal_idy,sal_dat from avt order by sal_idy", con, adOpenKeyset, adLockPessimistic
            If rschk.RecordCount > 0 Then rschk.MoveFirst
            rschk.filter = "sal_idy='" & Trim(RSINV(0)) & "' and sal_dat=#" & RSINV(1) & "#"
            
            If rschk.RecordCount < RSINV(3) Then
                Set lst = LVP.ListItems.Add(, , RSINV(0))
                'rsprd.Find "prd_idy='" & Trim(RSINV(2)) & "'"
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_idy,prd_nme from prd where prd_idy='" & Trim(RSINV(2)) & "'", con, adOpenKeyset, adLockPessimistic
                If Not rsprd.EOF Then lst.ListSubItems.Add , , rsprd(1)
                lst.ListSubItems.Add , , RSINV(3)
                lst.ListSubItems.Add , , RSINV(1)
                If rst_chk.State = 1 Then rst_chk.Close
                rst_chk.Open "select act_nme from act where act_idy='" & RSINV(4) & "'", con, adOpenKeyset, adLockPessimistic
                If rst_chk.EOF = False Then
                lst.ListSubItems.Add , , Trim(rst_chk(0))
                End If
                lst.ListSubItems.Add , , RSINV(5)
                If rst_chk.State = 1 Then rst_chk.Close
                rst_chk.Open "select smn_nme from smn where smn_idy='" & RSINV(6) & "'", con, adOpenKeyset, adLockPessimistic
                If rst_chk.EOF = False Then
                lst.ListSubItems.Add , , Trim(rst_chk(0))
                End If
                If Not rsprd.EOF Then lst.ListSubItems.Add , , Trim(rsprd(0)) & ""
                lst.ListSubItems.Add , , Trim(RSINV(4))
                lst.ListSubItems.Add , , Trim(RSINV(6))
                CHK_SAL = True
            End If
            RSINV.MoveNext
        Wend
        If CHK_SAL = True Then
            SSTab1.TabEnabled(3) = True
            SSTab1.Tab = 3
            LVP.SetFocus
        Else
       
        
        MsgBox "No Invoices Found.", vbOKOnly, "DAS Version 4.0"
        txtSalInv.SetFocus
        End If
    Else
        MsgBox "No Invoices Found.", vbOKOnly, "DAS Version 4.0"
        txtSalInv.SetFocus
    End If
'End If
'End If
If rst_chk.State = 1 Then rst_chk.Close
    Set rst_chk = Nothing
SAL_ERR:
'End If
End Sub

Public Function list_chk(A As Integer) As Boolean
Dim F1(10) As Integer
Dim F2(10) As Integer
Dim pn1(10) As String
list_chk = True
For p0 = 0 To Val(pn.Caption)
    F1(p0) = Mid(Ptno(p0), 1, Val(InStr(Ptno(p0), ":")) - 2)
    pn1(p0) = Mid(Ptno(p0), Val(InStr(Ptno(p0), ":")) + 2)
    
    For l1 = 1 To lv.ListItems.count
        If lv.ListItems(l1).Checked = True And lv.ListItems(l1).SubItems(1) = pn1(p0) Then
            F2(p0) = Val(F2(p0)) + 1
        End If
    Next
    If F1(p0) < F2(p0) Then
    list_chk = False
    GoTo AA
    End If
Next
AA:
End Function

Public Function CHK_SAL_IMP(SAL_ID As String, SUP_ID As String) As Boolean
Dim RST_INV As New ADODB.Recordset
Dim RST_INL As New ADODB.Recordset
Dim RST_PRD As New ADODB.Recordset
Dim RST_AVT As New ADODB.Recordset
If RST_INV.State = 1 Then RST_INV.Close
RST_INV.Open "SELECT sum(iNV.PRD_QTY),inv.prd_idy,prd.prd_nme,exc_inf FROM INV,prd WHERE INV_IDY='" & Trim(SAL_ID) & "' and inv.prd_idy=prd.prd_idy and exc_inf='Y' and inv.stk_typ in('Salable','Free') GROUP by inv.prd_idy", con, adOpenStatic
If RST_INV.EOF = False Then
    Do While RST_INV.EOF = False
    If RST_PRD.State = 1 Then RST_PRD.Close
    RST_PRD.Open "SELECT ALT_IDY FROM PRD WHERE PRD_IDY='" & Trim(RST_INV(1)) & "'", con, adOpenDynamic, adLockPessimistic
    If Trim(SUP_ID) = Trim(RST_PRD(0)) Then
        If RST_AVT.State = 1 Then sav_rst.Close
        RST_AVT.Open "select prd_idy from avt where sal_idy='" & Trim(SAL_ID) & "' and prd_idy='" & Trim(RST_INV(1)) & "'", con, adOpenKeyset, adLockPessimistic
        If RST_AVT.RecordCount < RST_INV(0) Then
            CHK_SAL_IMP = True
            Exit Do
        Else
            CHK_SAL_IMP = False
        End If
           
    End If
    RST_INV.MoveNext
    Loop
Else
    CHK_SAL_IMP = False
End If
End Function

Public Function CHK_RET(M As String) As Boolean

End Function

Public Sub CLEAR()
lv.ListItems.CLEAR
    txtSalInv = ""
    txtPartyname = ""
    sinvdat = Date$
    txtSmn = ""
    txtMarket = ""
    txtSMbl = ""
    txtSalInv.SetFocus
End Sub



