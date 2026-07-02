VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_IRT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Collection Entry"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "M_FRM_IRT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Bill Details"
      TabPicture(0)   =   "M_FRM_IRT.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "REC_LBL"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LBL_MSG"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Bnk_amt"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Bal_Amt"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Balance"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ProgressBar1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "DTPicker3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "MS"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "PST_LST"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "M_FLX_AMT"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "DTPicker4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Option1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_IRT.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LBL_NME"
      Tab(1).Control(1)=   "list5"
      Tab(1).Control(2)=   "PTY_LST"
      Tab(1).Control(3)=   "SMN_LST"
      Tab(1).Control(4)=   "ACT_LST"
      Tab(1).Control(5)=   "LDS_LST"
      Tab(1).Control(6)=   "VAN_LST"
      Tab(1).ControlCount=   7
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         TabIndex        =   47
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "Frame4"
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
         Height          =   300
         Left            =   2520
         TabIndex        =   37
         Top             =   4800
         Visible         =   0   'False
         Width           =   525
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   75
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3360
         MaxLength       =   70
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   255
         Left            =   960
         TabIndex        =   32
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   74383363
         CurrentDate     =   37687
      End
      Begin MSComctlLib.ListView VAN_LST 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6165
         View            =   3
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
            Text            =   "ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Salesman"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Route"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Reference"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Reco Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LDS_LST 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5530
         View            =   3
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
            Text            =   "ID"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   675
         Left            =   2040
         TabIndex        =   21
         Top             =   1455
         Width           =   3900
         Begin VB.TextBox M_INV_IDY 
            Appearance      =   0  'Flat
            CausesValidation=   0   'False
            Height          =   330
            Left            =   2640
            MaxLength       =   10
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "M_FRM_IRT.frx":047A
            Left            =   1080
            List            =   "M_FRM_IRT.frx":0490
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox VAN_IDY 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2640
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox LDS_IDY 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   2640
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox PTY_IDY 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2640
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox SMN_IDY 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2640
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Filter"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   300
            Width           =   825
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   4080
         TabIndex        =   16
         Top             =   360
         Width           =   4575
         Begin VB.TextBox BNK_IDY 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1215
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   607
            Width           =   1095
         End
         Begin VB.TextBox ACT_IDY 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1215
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   247
            Width           =   1095
         End
         Begin VB.Label BNK_LBL 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2400
            TabIndex        =   28
            Top             =   600
            Width           =   1965
         End
         Begin VB.Label ACT_LBL 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2400
            TabIndex        =   27
            Top             =   240
            Width           =   1965
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bank A/c ID"
            Height          =   195
            Left            =   165
            TabIndex        =   20
            Top             =   675
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cash A/c ID"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   315
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   2040
         TabIndex        =   11
         Top             =   360
         Width           =   1935
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   600
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   1250
            _ExtentX        =   2196
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74383363
            CurrentDate     =   37258
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   600
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   600
            Width           =   1250
            _ExtentX        =   2196
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74383363
            CurrentDate     =   37258
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            Height          =   195
            Left            =   270
            TabIndex        =   15
            Top             =   675
            Width           =   195
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   315
            Width           =   345
         End
      End
      Begin MSComctlLib.ListView ACT_LST 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6165
         View            =   3
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
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox M_FLX_AMT 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   7080
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox PST_LST 
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
         ItemData        =   "M_FRM_IRT.frx":04D3
         Left            =   6120
         List            =   "M_FRM_IRT.frx":04E0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   4200
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ListView SMN_LST 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6165
         View            =   3
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
            Text            =   "Salesman Name"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   2595
         Left            =   100
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2160
         Width           =   8600
         _ExtentX        =   15161
         _ExtentY        =   4577
         _Version        =   393216
         Cols            =   14
         FixedCols       =   2
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         WordWrap        =   -1  'True
         FocusRect       =   0
         HighLight       =   2
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
      Begin MSComctlLib.ListView PTY_LST 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6165
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Market Name"
            Object.Width           =   6068
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   330
         Left            =   750
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   600
         Width           =   1250
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74383363
         CurrentDate     =   37258
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   3375
         Begin VB.CheckBox ALL_CHK 
            Caption         =   "&All"
            Height          =   375
            Left            =   1400
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   0
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView list5 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   44
         Top             =   1080
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6165
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
            Text            =   "Invoice  Number"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Retailer Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Place"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   6120
         TabIndex        =   29
         Top             =   1695
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Balance 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
         Height          =   195
         Left            =   3960
         TabIndex        =   49
         Top             =   4845
         Width           =   585
      End
      Begin VB.Label Bal_Amt 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4680
         TabIndex        =   48
         Top             =   4845
         Width           =   75
      End
      Begin VB.Label Bnk_amt 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7725
         TabIndex        =   46
         Top             =   4845
         Width           =   75
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bank"
         Height          =   195
         Left            =   7305
         TabIndex        =   45
         Top             =   4845
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Receipt Date"
         Height          =   420
         Left            =   120
         TabIndex        =   35
         Top             =   555
         Width           =   555
      End
      Begin VB.Label LBL_MSG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Actual Amount Received"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   4845
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cash"
         Height          =   195
         Left            =   5850
         TabIndex        =   10
         Top             =   4845
         Width           =   360
      End
      Begin VB.Label REC_LBL 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6375
         TabIndex        =   9
         Top             =   4845
         Width           =   75
      End
      Begin VB.Label LBL_NME 
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
         Left            =   -74760
         TabIndex        =   4
         Top             =   720
         Width           =   8295
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
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
            Picture         =   "M_FRM_IRT.frx":04ED
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IRT.frx":0A31
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IRT.frx":0DF9
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IRT.frx":114D
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IRT.frx":1785
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IRT.frx":1AD9
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IRT.frx":1F39
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IRT.frx":222D
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IRT.frx":2639
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IRT.frx":2745
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IRT.frx":2A99
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IRT.frx":2EAD
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
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
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ed"
            Object.ToolTipText     =   "View (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "idc"
            Object.ToolTipText     =   "Click - Generate Invoice Collection History"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "pre"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "d"
            Object.ToolTipText     =   "Delete (Ctrl+D)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "&Update Entries"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7440
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   0
         Width           =   1335
      End
   End
End
Attribute VB_Name = "M_FRM_IRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean, CTR As Long, COM_CTR As Long
Private Sub ACT_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Cash Account"
End Sub

Private Sub ACT_IDY_KeyPress(K As Integer)
If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 13 Then
    Dim REC_MAR As New ADODB.Recordset
    Dim REC_SET As New ADODB.Recordset
    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where ACT_IDY = '" & Sincrement(ACT_IDY) & "' AND GRP_CLF in ('Bank Book','Cash Book') AND DBT_TYP LIKE 'Y'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        ACT_LBL = REC_SET("ACT_NME")
        ACT_IDY = Sincrement(ACT_IDY)
        BNK_IDY.SetFocus
    Else
        V_ACT_LST
    End If
End If
End Sub

Private Sub ACT_LST_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If ACT_LST.SortOrder = lvwAscending Then
    ACT_LST.SortOrder = lvwDescending
ElseIf ACT_LST.SortOrder = lvwDescending Then
    ACT_LST.SortOrder = lvwAscending
End If
ACT_LST.SortKey = CH.Index - 1
End Sub

Private Sub ACT_LST_DblClick()
ACT_LST_KeyPress 13
End Sub

Private Sub ACT_LST_KeyPress(K As Integer)
If K = 13 Then
    If ACT_LST.ListItems.count > 0 Then
        If LBL_NME = "Select Cash Account" Then
            ACT_IDY = ACT_LST.SelectedItem.ListSubItems.Item(1)
            ACT_LBL = ACT_LST.SelectedItem
            ACT_LST.Visible = False
            SSTab1.Tab = 0
            BNK_IDY.SetFocus
        ElseIf LBL_NME = "Select Bank Account" Then
            BNK_IDY = ACT_LST.SelectedItem.ListSubItems.Item(1)
            BNK_LBL = ACT_LST.SelectedItem
            ACT_LST.Visible = False
            SSTab1.Tab = 0
            If ALL_CHK.Enabled = True Then
                ALL_CHK.SetFocus
            Else
                MS.SetFocus
            End If
        End If
    End If
Else
End If
End Sub

Private Sub ALL_CHK_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check To Load All Invoices"
End Sub

Private Sub ALL_CHK_KeyPress(K As Integer)
If K = 13 Then
    If ALL_CHK.Value = 0 Then
        ALL_CHK.Enabled = False
        Frame3.Enabled = True
        Combo1.Enabled = True
        Combo1.SetFocus
        Combo1.ListIndex = 4
    ElseIf ALL_CHK.Value = 1 Then
        ALL_CHK.Enabled = False
        Frame3.Enabled = False
        displaydetails
        For A = 1 To MS.Rows - 1
            If MS.TextMatrix(1, 0) = "" Then
                MS.SetFocus
            Else
                MS.Col = 5
            End If
        Next
    End If

'    If ALL_CHK.value = 1 Then
'        ALL_CHK.Enabled = False
'        Frame3.Enabled = False
'        displaydetails
'        For A = 1 To MS.Rows - 1
'            If MS.TextMatrix(1, 0) = "" Then
'                MS.SetFocus
'            Else
'                MS.Col = 5
'            End If
'        Next
'    ElseIf ALL_CHK.value = 0 Then
'        ALL_CHK.Enabled = False
'        Frame3.Enabled = True
'
'    End If
ElseIf K = 38 Then
    DTPicker2.SetFocus
ElseIf K = 32 Then
    Frame3.Enabled = True
    ALL_CHK.Enabled = False
    Combo1.Enabled = True
    Combo1.SetFocus
    Combo1.ListIndex = 4
    displaydetails
    For A = 1 To MS.Rows - 1
        If MS.TextMatrix(1, 0) = "" Then
            MS.SetFocus
        Else
            MS.Col = 5
        End If
    Next
End If
End Sub

Private Sub BNK_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Bank Account"
End Sub

Private Sub BNK_IDY_Keypress(K As Integer)
If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
            
If K = 39 Or K = 41 Or K = 37 Then K = 0
Dim rsact As New ADODB.Recordset
If K > 96 And K < 123 Then K = K - 32
If K = 13 Then
    rsact.Open "Select * from act where grp_clf like 'Bank Book' and act_idy like '" & Sincrement(BNK_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If rsact.RecordCount > 0 Then
        BNK_IDY = Sincrement(BNK_IDY)
        If ALL_CHK.Enabled = True Then
            ALL_CHK.Value = 0 'Added by PNR as per Prince Chopra feedback
            ALL_CHK.SetFocus
        Else
            MS.SetFocus
        End If
    Else
        V_bnk_LST
    End If
End If
End Sub

Private Sub Combo1_keypress(K As Integer)
If K = 13 Then
    If Combo1.ListIndex <> 0 Then Combo1.Enabled = False
    If Combo1.ListIndex = 1 Then
        SMN_IDY.Visible = True
        SMN_IDY.Enabled = True
        LDS_IDY.Visible = False
        VAN_IDY.Visible = False
        M_INV_IDY.Visible = False
        PTY_IDY.Visible = False
        SMN_IDY.SetFocus
    ElseIf Combo1.ListIndex = 2 Then
        SMN_IDY.Visible = False
        LDS_IDY.Visible = False
        M_INV_IDY.Visible = False
        VAN_IDY.Visible = False
        PTY_IDY.Visible = True
        PTY_IDY.Enabled = True
        PTY_IDY.SetFocus
    ElseIf Combo1.ListIndex = 3 Then
        SMN_IDY.Visible = False
        LDS_IDY.Visible = True
        LDS_IDY.Enabled = True
        VAN_IDY.Visible = False
        M_INV_IDY.Visible = False
        PTY_IDY.Visible = False
        LDS_IDY.SetFocus
    ElseIf Combo1.ListIndex = 4 Then
        SMN_IDY.Visible = False
        M_INV_IDY.Visible = False
        LDS_IDY.Visible = False
        VAN_IDY.Visible = True
        VAN_IDY.Enabled = True
        PTY_IDY.Visible = False
        VAN_IDY.SetFocus
    ElseIf Combo1.ListIndex = 5 Then
        SMN_IDY.Visible = False
        M_INV_IDY.Visible = True
        M_INV_IDY.Enabled = True
        LDS_IDY.Visible = False
        VAN_IDY.Visible = False
        PTY_IDY.Visible = False
        M_INV_IDY.SetFocus
    ElseIf Combo1.ListIndex = 0 Then
        MsgBox "Press Down Arrow To Select Filter !", vbExclamation
        Combo1.SetFocus
    End If
End If
End Sub

Private Sub V_INV_LST()
If CHECKTB("act.dbf,inl.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select inl.inv_idy,inl.inv_dat,act.act_nme,act.act_ad4 from act,inl where act_nme like '" & Trim(M_INV_IDY) & "%' and inl.cst_idy=act.act_idy and inl.txn_typ ='SAL'", con, adOpenKeyset, adLockPessimistic
list5.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list5.ListItems.Add(, , Trim(REC_SET(0)))
        ls.ListSubItems.Add , , Format(REC_SET(1), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_AD4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    LBL_NME = "Select Invoice"
    list5.Visible = True
    list5.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Invoice Number Not Found!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        M_INV_IDY.Enabled = True
        SendKeys "{HOME}+{END}"
        M_INV_IDY.SetFocus
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub

Private Sub Command1_Click()
Command1.Enabled = False
For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 5)) = 0 Then
        MS.TextMatrix(A, 5) = convert(Val(MS.TextMatrix(A, 4)))
        MS.TextMatrix(A, 4) = "0.00"
    End If
Next
Total
MS.SetFocus
End Sub

Private Sub DTPicker1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select 'From' Date"
End Sub

Private Sub DTPicker3_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    If DTPicker1.Enabled = True Then
        DTPicker1.SetFocus
    Else
        MS.SetFocus
    End If
End If
End Sub

Private Sub DTPicker1_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select 'To' Date"
End Sub

Private Sub DTPicker2_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    ACT_IDY.SetFocus
End If
If K = 38 And s = 1 Then DTPicker1.SetFocus
End Sub

Private Sub DTPicker2_LostFocus()
If DTPicker2 < DTPicker1 Then
    MsgBox " 'To' Date Should Be Greater Than 'From' Date", vbExclamation, "DAS Version 4.0"
    DTPicker2.SetFocus
    Exit Sub
End If
End Sub

Private Sub DTPicker4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MS.Text = Format(DTPicker4, "DD/MMM/YYYY")
    DTPicker4.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub LIST5_KeyPress(K As Integer)
If K = 13 Then
    M_INV_IDY = list5.SelectedItem
    list5.Visible = False
    SSTab1.Tab = 0
    displaydetails
End If
End Sub

Private Sub M_FLX_AMT_GotFocus()
'Total
End Sub

Private Sub M_INV_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 47 And K < 58 Or K = 8 Or K = 13 Then
ElseIf K > 96 And K < 123 Then
K = K - 32
ElseIf K > 64 And K < 91 Then
Else
K = 0
End If
If M_INV_IDY = "" And K = 27 Then Unload Me
Set Form = Nothing
End Sub
Private Sub M_INV_IDY_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    If Len(Trim(M_INV_IDY)) = 0 Then
        V_INV_LST
    Else
        If check("inl", "inv_idy", Trim(M_INV_IDY)) = True Then
            M_INV_IDY = Trim(M_INV_IDY)
            displaydetails
        Else
            V_INV_LST
        End If
    End If
End If
End Sub

Private Sub MS_Scroll()
If DTPicker4.Visible = True Then
    DTPicker4.Visible = False
    MS.SetFocus
ElseIf Text1.Visible = True Then
    Text1 = ""
    Text1.Visible = False
    MS.SetFocus
ElseIf PST_LST.Visible = True Then
    PST_LST.Visible = False
    MS.SetFocus
ElseIf M_FLX_AMT.Visible = True Then
    M_FLX_AMT = ""
    M_FLX_AMT.Visible = False
    MS.SetFocus
End If
End Sub
Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 37 Then K = 0
If K = 13 Then
    MS.Text = Text1
    Text1 = ""
    Text1.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub LDS_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Load Sheet"
End Sub

Private Sub LDS_IDY_KeyPress(K As Integer)
If CHECKTB("lds.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
If K = 13 Then
    If check("lds", "lds_idy", "LDS" & Tincrement(LDS_IDY)) = True Then
        PTY_IDY = ""
        SMN_IDY = ""
        VAN_IDY = ""
        ALL_CHK.Value = 0
        LDS_IDY = "LDS" & Tincrement(LDS_IDY)
        displaydetails
    Else
        V_LDS_LST
    End If
End If
End Sub
Private Sub VAN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Van Ref."
End Sub

Private Sub VAN_IDY_KeyPress(K As Integer)
If CHECKTB("VAN.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
If K = 13 Then
    If check("VAN", "INV_IDY", "RSI" & Tincrement(VAN_IDY)) = True Then
        PTY_IDY = ""
        SMN_IDY = ""
        LDS_IDY = ""
        ALL_CHK.Value = 0
        VAN_IDY.Enabled = False
        VAN_IDY = "RSI" & Tincrement(VAN_IDY)
        displaydetails
    Else
        V_RSI_LST
    End If
End If
End Sub

Private Sub LDS_LST_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    LDS_IDY = LDS_LST.SelectedItem
    LDS_LST.Visible = False
    PTY_IDY = ""
    ALL_CHK.Value = 0
    SMN_IDY = ""
    SSTab1.Tab = 0
    LDS_IDY.Enabled = False
    displaydetails
End If
End Sub

Private Sub MS_EnterCell()
If MS.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Customer Name"
ElseIf MS.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Invoice ID"
ElseIf MS.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "Invoice Amount"
ElseIf MS.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Received Amount"
ElseIf MS.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Balance Amount"
ElseIf MS.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Receiving Amount"
ElseIf MS.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "Select 'C' For Cash Account or 'B' For Bank Account"
ElseIf MS.Col = 8 Then
    M_MDI.STS_BAR.Panels(1).Text = "Invoice Date"
ElseIf MS.Col = 7 Then
    M_MDI.STS_BAR.Panels(1).Text = "Select 'Y' To Post Receipts, 'N' Not To Post Receipts or 'D' To Delete The Record"
ElseIf MS.Col = 9 Then
    M_MDI.STS_BAR.Panels(1).Text = "Due Date"
ElseIf MS.Col = 10 Then
    M_MDI.STS_BAR.Panels(1).Text = "Due Days"
End If
End Sub

Private Sub MS_GotFocus()
MS.HighLight = flexHighlightAlways
End Sub

Private Sub MS_KeyPress(K As Integer)
If (MS.Col = 6 Or MS.Col = 5) And MS.Row > 0 Then
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
        If K = vbKeySubtract Then
        Else
            M_FLX_AMT = Chr(K)
            M_FLX_AMT.SelStart = Len(M_FLX_AMT)
        End If
    End If
End If
If MS.Row > 0 Then
    If MS.Col = 12 Then
        With DTPicker4
        .Visible = True
        .Top = MS.Top + MS.CellTop
        .Left = MS.CellLeft + MS.Left
        .Width = MS.CellWidth
        .Height = MS.CellHeight
        .Visible = True
        .SetFocus
        End With
    End If
    If K <> 13 Then
        If MS.Col = 13 Then
            Text1.Visible = True
            Text1.Top = MS.Top + MS.CellTop
            Text1.Left = MS.CellLeft + MS.Left
            Text1.Width = MS.CellWidth
            Text1.Height = MS.CellHeight
            
            If K <> 8 Then Text1 = Chr(K)
            Text1.Visible = True
            Text1.SetFocus
            Text1.SelStart = Len(Text1)
        End If
    End If
End If
If MS.Col = 7 And MS.Row > 0 Then
    If Chr(K) = "y" Or Chr(K) = "Y" Then
        MS.Text = "Y"
    ElseIf Chr(K) = "N" Or Chr(K) = "n" Then
        MS.Text = "N"
    ElseIf K = vbKeySpace Then
        If MS.Text = "Y" Then
            MS.Text = "N"
        ElseIf MS.Text = "N" Then
            MS.Text = "Y"
        End If
    End If
End If
End Sub

Private Sub MS_LeaveCell()
If M_FLX_AMT.Visible = True Then M_FLX_AMT.Visible = False
End Sub

Public Sub displaydetails()
    Dim CHK_CUR As New ADODB.Recordset
    If CHK_CUR.State = 1 Then CHK_CUR.Close
    CHK_CUR.Open "SELECT * FROM CUR WHERE INV_IDY = '" & Trim(VAN_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If CHK_CUR.RecordCount > 0 Then
        If CHK_CUR("CHG_IDY") = 1 Then
            MsgBox "Reconciliation Pending for NetSales and Stock for " & Trim(VAN_IDY) & ", Collection Entry Not Possible at this Moment", vbInformation, "DAS Version 4.0"
            CANCELRECORD
            Exit Sub
        End If
    End If

If MS.Rows < 3 Then
    If CHECKTB("vch.dbf,inl.dbf,act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If
    
    
    Dim REC_SET As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim RS1 As New ADODB.Recordset
    If REC_SET.State = 1 Then REC_SET.Close
    
    If Combo1.ListIndex = 4 Then
        REC_SET.Open "SELECT SUM(TXN_AMT),INV_IDY FROM VCH WHERE INV_TYP = 'SAL' AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE NET_AMT > 0 AND TXN_TYP = 'SAL' AND ORD_IDY LIKE '" & Trim(VAN_IDY) & "%')GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
    ElseIf Combo1.ListIndex = 3 Then
        REC_SET.Open "SELECT SUM(TXN_AMT),INV_IDY FROM VCH WHERE INV_TYP = 'SAL' AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE NET_AMT > 0 AND TXN_TYP = 'SAL' AND DCH_IDY LIKE '" & Trim(LDS_IDY) & "%')GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
    ElseIf Combo1.ListIndex = 0 Then
        REC_SET.Open "SELECT SUM(TXN_AMT),INV_IDY FROM VCH WHERE INV_TYP = 'SAL' AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE NET_AMT > 0 AND TXN_TYP = 'SAL' AND CST_IDY LIKE '" & Trim(PTY_IDY) & "%' AND SMN_IDY LIKE '" & Trim(SMN_IDY) & "%' AND DCH_IDY LIKE '" & Trim(LDS_IDY) & "%' AND ORD_IDY LIKE '" & Trim(VAN_IDY) & "%' AND INV_DAT >= ctod('" & Format(DTPicker1, "mm/DD/yyyy") & "') AND INV_DAT <= ctod('" & Format(DTPicker2, "mm/DD/yyyy") & "')) GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
    ElseIf Combo1.ListIndex = 5 Then
        REC_SET.Open "SELECT SUM(TXN_AMT),INV_IDY FROM VCH WHERE INV_TYP = 'SAL' AND INV_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    Else
        REC_SET.Open "SELECT SUM(TXN_AMT),INV_IDY FROM VCH WHERE INV_TYP = 'SAL' AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE NET_AMT > 0 AND TXN_TYP = 'SAL' AND CST_IDY LIKE '" & Trim(PTY_IDY) & "%' AND SMN_IDY LIKE '" & Trim(SMN_IDY) & "%' AND DCH_IDY LIKE '" & Trim(LDS_IDY) & "%' AND ORD_IDY LIKE '" & Trim(VAN_IDY) & "%' AND INV_DAT >= ctod('" & Format(DTPicker1, "mm/DD/yyyy") & "') AND INV_DAT <= ctod('" & Format(DTPicker2, "mm/DD/yyyy") & "')) GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
    End If
    If RS1.State = 1 Then rs.Close
    If Combo1.ListIndex = 5 Then
        RS1.Open "SELECT * FROM INL WHERE NET_AMT > 0 AND TXN_TYP = 'SAL' AND INL.inv_IDY = '" & Trim(M_INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    Else
        RS1.Open "SELECT * FROM INL WHERE NET_AMT > 0 AND TXN_TYP = 'SAL' AND INL.CST_IDY LIKE '" & Trim(PTY_IDY) & "%' AND INL.SMN_IDY LIKE '" & Trim(SMN_IDY) & "%' AND INL.DCH_IDY LIKE '" & Trim(LDS_IDY) & "%' AND INV_DAT >= ctod('" & Format(DTPicker1, "mm/DD/yyyy") & "') AND INV_DAT <= ctod('" & Format(DTPicker2, "mm/DD/yyyy") & "') AND INL.ORD_IDY LIKE '" & Trim(VAN_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
    End If
    If rs.State = 1 Then rs.Close
    COM_CTR = 0
    'Label5.Visible = False
    Me.MousePointer = vbHourglass
''    PROC.Visible = True
'    PROC.Refresh
    LBL_MSG.Visible = False
    CTR = REC_SET.RecordCount + RS1.RecordCount
    rs.Open "SELECT * FROM ACT", con, adOpenKeyset, adLockPessimistic
    con.Execute "update IRT set CHK_FLG = 'Y' "
    MS.Rows = RS1.RecordCount + 1
    'PROC.Caption = "Loading Records... Please Wait"
'    PROC.Refresh
    ProgressBar1.Visible = True
    DoEvents
'    Frame4.Visible = True
'    Frame4.BackColor = &H80000000
    For A = 1 To RS1.RecordCount
        If rs.BOF = False Then rs.MoveFirst
        rs.Find "ACT_IDY = '" & Trim(RS1("CST_IDY")) & "'"
        If rs.EOF = False Then MS.TextMatrix(A, 0) = Trim(rs("ACT_NME"))
        MS.TextMatrix(A, 1) = Trim(RS1("INV_IDY"))
        MS.TextMatrix(A, 8) = Format(RS1("INV_DAT"), "DD/MMM/YYYY")
        MS.TextMatrix(A, 9) = Format(RS1("DUE_DAT"), "DD/MMM/YYYY") & ""
        MS.TextMatrix(A, 2) = convert(RS1("NET_AMT"))
        MS.TextMatrix(A, 3) = "0.00"
        MS.TextMatrix(A, 4) = convert(Val(MS.TextMatrix(A, 2)))
        MS.TextMatrix(A, 5) = "0.00"
        MS.TextMatrix(A, 6) = "0.00"
        MS.TextMatrix(A, 7) = "Y"
        MS.TextMatrix(A, 10) = DateDiff("d", RS1("INV_DAT"), Date)
        MS.TextMatrix(A, 11) = RS1("CST_IDY")
        MS.TextMatrix(A, 12) = Format(Date, "DD/MMM/YYYY")
        REC_LBL = Val(REC_LBL) + Val(MS.TextMatrix(A, 5))
        
        
        
        incpcb
        RS1.MoveNext
    Next
    
    For A = 1 To REC_SET.RecordCount
        For b = 1 To MS.Rows - 1
            If Trim(MS.TextMatrix(b, 1)) = Trim(REC_SET!INV_IDY) Then
                If Val(MS.TextMatrix(b, 2)) = Val(REC_SET(0) & "") Then
                    If MS.Rows = 2 Then
                        MS.Rows = 1
                    Else
                        MS.RemoveItem b
                    End If
                Else
                    MS.TextMatrix(b, 3) = convert(Val(REC_SET(0) & ""))
                    MS.TextMatrix(b, 4) = convert(Val(MS.TextMatrix(b, 2)) - Val(REC_SET(0) & ""))
                End If
                Exit For
            End If
        Next
        incpcb
        REC_SET.MoveNext
    Next
    DoEvents

    ProgressBar1.Visible = False
    Frame4.Visible = False
'    PROC.Visible = False
    Label5.Visible = True
    If Combo1.ListIndex = 3 Or Combo1.ListIndex = 4 Then
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    End If
    Frame3.Enabled = False
    SMN_IDY.Enabled = False
    PTY_IDY.Enabled = False
    LDS_IDY.Enabled = False
    VAN_IDY.Enabled = False
    Command1.Enabled = True
    DoEvents
    If MS.Rows > 1 Then LBL_MSG.Visible = True
    Me.MousePointer = vbNormal
End If
If Combo1.ListIndex = 5 Then 'Added by PKB
    If Option1.Value = True Then
        If MsgBox("Generate Collection History ? ", vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
            SAL_QRY = "VCH->INV_IDY = '" & Trim(M_INV_IDY.Text) & "'"
            SAL_QRY = SAL_QRY & ".AND. VCH->INV_TYP = 'SAL'" 'Changed by PKB
            OpenReport CURDIR & "\REP_LIB.RP1", "Collection Breakup Invoice", 2, SAL_QRY, 6
            CANCELRECORD
            Exit Sub
'        Else
'            CANCELRECORD
'            Exit Sub
        End If
    Else
    End If
End If
If MS.Rows = 1 Then
    MsgBox "Outstanding Invoice Information Not Available!", vbExclamation, "DAS Version 4.0"
    CANCELRECORD
    DTPicker1.SetFocus
Else
    MS.SetFocus
    MS.Row = 1
    MS.Col = 5
    Total
    MS.SetFocus
    MS.Row = 1
End If
End Sub

Public Sub incpcb()
COM_CTR = COM_CTR + 1
ProgressBar1.Value = (100 / CTR) * COM_CTR
Label7.Caption = Int((100 / CTR) * COM_CTR) & "%"
If Int((100 / CTR) * COM_CTR) = 55 Then
'    Frame4.BackColor = &H800000
End If
Label7.Refresh
End Sub
Private Sub M_FLX_AMT_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    MS.Text = amt(M_FLX_AMT)
    M_FLX_AMT.Visible = False
    If MS.Col = 5 Or MS.Col = 6 Then
        For A = 1 To MS.Rows - 1
            MS.TextMatrix(A, 4) = convert(Val(Val(MS.TextMatrix(A, 2)) - (Val(MS.TextMatrix(A, 6)) + Val(MS.TextMatrix(A, 3)) + Val(MS.TextMatrix(A, 5)))))
            If Val(MS.TextMatrix(A, 5)) > 0 Or Val(MS.TextMatrix(A, 6)) > 0 Then
                MS.TextMatrix(A, 7) = "Y"
            Else
                MS.TextMatrix(A, 7) = "N"
            End If

        Next
    End If
    If MS.Row = MS.Rows - 1 Then
    Else
        MS.Row = MS.Row + 1
    End If
    Total
    MS.SetFocus
End If
End Sub

Private Sub M_FLX_AMT_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 13 Or K = 8 Then
Else
    K = 0
End If
End Sub

Private Sub Form_Activate()
checkdongle
SETYEAR M_FRM_IRT
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
PTY_LST.Visible = False
SMN_LST.Visible = False
DTPicker3.SetFocus
End Sub

Private Sub V_SMN_LST()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from SMN where SMN_NME like '" & Trim(SMN_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
SMN_LST.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = SMN_LST.ListItems.Add(, , Trim(REC_SET("SMN_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("SMN_idy"))
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    LBL_NME.Caption = "Select Salesman"
    SMN_LST.Visible = True
    SMN_LST.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        SMN_IDY.SetFocus
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub

Private Sub V_PTY_LST()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_MAR As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from act where act_NME like '" & Trim(PTY_IDY) & "%' and grp_clf like 'Debtor' AND DBT_TYP LIKE 'Y'", con, adOpenKeyset, adLockPessimistic
PTY_LST.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = PTY_LST.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    LBL_NME.Caption = "Select Account"
    SMN_LST.Visible = False
    ACT_LST.Visible = False
    LDS_LST.Visible = False
    PTY_LST.Visible = True
    PTY_LST.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        PTY_IDY.SetFocus
    End If
End If
REC_SET.Close
'REC_MAR.Close
Set REC_SET = Nothing
Set REC_MAR = Nothing
End Sub

Private Sub V_ACT_LST()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_MAR As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from act where ACT_NME like '" & Trim(ACT_IDY) & "%' AND GRP_CLF in ('Cash Book') AND DBT_TYP LIKE 'Y'", con, adOpenKeyset, adLockPessimistic
ACT_LST.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = ACT_LST.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    PTY_LST.Visible = False
    SMN_LST.Visible = False
    ACT_LST.Visible = True
    LBL_NME.Caption = "Select Cash Account"
    ACT_LST.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        ACT_IDY.SetFocus
    End If
End If
'REC_MAR.Close
REC_SET.Close
Set REC_MAR = Nothing
Set REC_SET = Nothing
End Sub

Private Sub V_bnk_LST()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from act where ACT_NME like '" & Trim(BNK_IDY) & "%' AND GRP_CLF in ('Bank Book') AND DBT_TYP LIKE 'Y'", con, adOpenKeyset, adLockPessimistic
ACT_LST.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = ACT_LST.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    PTY_LST.Visible = False
    SMN_LST.Visible = False
    ACT_LST.Visible = True
    LBL_NME.Caption = "Select Bank Account"
    ACT_LST.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        BNK_IDY.SetFocus
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub
Private Sub V_RSI_LST()
If CHECKTB("van.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
qry = ""
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
'REC_SET.Open "SELECT * FROM VAN WHERE INV_IDY LIKE '" & Trim(VAN_IDY) & "%' GROUP BY INV_IDY", CON, adOpenKeyset, adlockpessimistic
REC_SET.Open "select VAN.INV_IDY,INV_DAT,SMN_NME,SMN_AD4,VAN_CLS,ROU_IDY,MNL_REF,CUR.CHG_IDY from SMN,VAN,CUR where SMN_nme like '" & Trim(M_INV_IDY) & "%' and VAN.smn_idy=smn.smn_idy AND VAN.INV_IDY = CUR.INV_IDY GROUP BY VAN.INV_IDY ", con, adOpenKeyset, adLockPessimistic 'CHANGED BY PNR
VAN_LST.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        If REC_SET("CHG_IDY") = 0 Then
            REC_FLG = "ReconCiled"
        Else
            REC_FLG = "RecoPendng"
        End If
        Set ls = VAN_LST.ListItems.Add(, , Trim(REC_SET("INV_IDY")))
        ls.ListSubItems.Add , , Format(REC_SET("INV_DAT"), "DD/MM/YYYY")
        ls.ListSubItems.Add , , REC_SET("SMN_NME")
        ls.ListSubItems.Add , , REC_SET("ROU_IDY")
        ls.ListSubItems.Add , , REC_SET("MNL_REF")
        ls.ListSubItems.Add , , REC_FLG
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    VAN_LST.Visible = True
    VAN_LST.SetFocus
    LBL_NME.Caption = "Select Van Reference"
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Van Reference Not Found!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        VAN_IDY.SetFocus
    End If
End If
REC_SET.Close
Set REC_SET = Nothing
End Sub
Private Sub V_LDS_LST()
If CHECKTB("lds.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim REC_MAR As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from LDS WHERE LDS_IDY LIKE '" & Trim(LDS_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
LDS_LST.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = LDS_LST.ListItems.Add(, , Trim(REC_SET("LDS_IDY")))
        ls.ListSubItems.Add , , Format(REC_SET("LDS_DAT"), "DD/MM/YYYY")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    LDS_LST.Visible = True
    LDS_LST.SetFocus
    LBL_NME.Caption = "Select Load Sheet"
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Load Sheet Not Found!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        LDS_IDY.SetFocus
    End If
End If
REC_SET.Close
REC_MAR.Close
Set REC_MAR = Nothing
Set REC_SET = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 17 Then KeyCode = 0
If KeyCode = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If PTY_LST.Visible = False And SMN_LST.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then newrec
    If KeyCode = 80 And Shift = 2 Then Printrecord
    If KeyCode = 83 And Shift = 2 Then saverecord
    If KeyCode = 88 And Shift = 2 Then Unload Me
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
        If M_FRM_IRT.ActiveControl.name = "pty_IDY" Then
            V_PTY_LST
        ElseIf M_FRM_IRT.ActiveControl.name = "Smn_IDY" Then
            V_SMN_LST
        ElseIf M_FRM_IRT.ActiveControl.name = "ACT_IDY" Then
            V_ACT_LST
        ElseIf M_FRM_IRT.ActiveControl.name = "M_INV_IDY" Then
            V_INV_LST
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_IRT.ActiveControl Is TextBox Then M_FRM_IRT.ActiveControl.Text = FUNKEY(KeyCode)
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If ACT_LST.Visible = True And LBL_NME = "Select Cash Account" Then
        SSTab1.Tab = 0
        ACT_LST.Visible = False
        ACT_IDY.SetFocus
    ElseIf ACT_LST.Visible = True And LBL_NME = "Select Bank Account" Then
        SSTab1.Tab = 0
        ACT_LST.Visible = False
        BNK_IDY.SetFocus
    ElseIf PTY_LST.Visible = True Then
        SSTab1.Tab = 0
        PTY_LST.Visible = False
        PTY_IDY.SetFocus
    ElseIf SMN_LST.Visible = True Then
        SSTab1.Tab = 0
        SMN_LST.Visible = False
        SMN_IDY.SetFocus
    ElseIf LDS_LST.Visible = True Then
        SSTab1.Tab = 0
        LDS_LST.Visible = False
        LDS_IDY.SetFocus
    ElseIf VAN_LST.Visible = True Then
        SSTab1.Tab = 0
        VAN_LST.Visible = False
        VAN_IDY.SetFocus
    ElseIf list5.Visible = True Then
        SSTab1.Tab = 0
        list5.Visible = False
        M_INV_IDY.SetFocus
    Else
        If M_FRM_IRT.ActiveControl.name = "DTPicker1" Then
            Unload Me
            Set Form = Nothing
        Else
            saverecord
        End If
    End If
End If
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
MS.TextMatrix(0, 0) = "Customer  "
MS.TextMatrix(0, 1) = "Invoice ID"
MS.TextMatrix(0, 2) = "Invoice Amt"
MS.TextMatrix(0, 3) = "Received Amt"
MS.TextMatrix(0, 4) = "Balance"
MS.TextMatrix(0, 5) = "Cash"
MS.Row = 0
MS.Col = 5
MS.CellForeColor = vbBlue
MS.TextMatrix(0, 6) = "Bank"
MS.Row = 0
MS.Col = 6
MS.CellForeColor = vbBlue
MS.TextMatrix(0, 7) = "Post(Y/N)"
MS.TextMatrix(0, 8) = "Inv. Date"
MS.ColAlignment(8) = 1
MS.TextMatrix(0, 9) = "Due Date"
MS.ColAlignment(9) = 1
MS.TextMatrix(0, 10) = "Due Days"
MS.TextMatrix(0, 12) = "Dep.Date"
MS.ColAlignment(12) = 1
MS.Row = 0
MS.Col = 12
MS.CellForeColor = vbBlue
MS.TextMatrix(0, 13) = "Remarks"
MS.Row = 0
MS.Col = 13
MS.CellForeColor = vbBlue
MS.ColWidth(0) = 2500
MS.ColWidth(1) = 1050
MS.ColAlignment(2) = 7
MS.ColAlignment(3) = 7
MS.ColAlignment(4) = 7
MS.ColAlignment(5) = 7
MS.ColAlignment(6) = 7
MS.ColAlignment(10) = 7
MS.ColWidth(2) = 1000
MS.ColWidth(3) = 1200
MS.ColWidth(4) = 1100
MS.ColWidth(5) = 1200
MS.ColWidth(6) = 1200
MS.ColWidth(7) = 0
MS.ColWidth(8) = 1000
MS.ColWidth(9) = 1000
MS.ColWidth(10) = 1000
MS.ColWidth(11) = 0
MS.ColWidth(12) = 1200
MS.ColWidth(13) = 6000 'Changed by PKB
REC_LBL.Caption = "0.00"
Bnk_amt.Caption = "0.00"
MOD_REC = False
End Sub

Private Sub PST_LST_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim REC_SET As New ADODB.Recordset
    PST_LST.Visible = False
    MS.Text = PST_LST
    PST_LST.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub PTY_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Party"
End Sub

Private Sub pty_IDY_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ls As ListItem
Dim REC_MAR As New ADODB.Recordset
Dim rs As New ADODB.Recordset
If KeyCode = 13 Then
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
rs.Open "select * from act where ACT_IDY like '" & Sincrement(Trim(PTY_IDY)) & "' and grp_clf like 'Debtor' AND DBT_TYP LIKE 'Y'", con, adOpenStatic
PTY_LST.ListItems.CLEAR
If rs.RecordCount > 0 Then
    PTY_IDY = Sincrement(Trim(PTY_IDY))
    ALL_CHK.Value = 0
    
    SMN_IDY = ""
    
    VAN_IDY = ""
    
    LDS_IDY = ""
    displaydetails
Else
    V_PTY_LST
End If
End If
End Sub

Private Sub pty_IDY_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 13 Then MS.SetFocus
End Sub

Private Sub PTY_LST_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If PTY_LST.SortOrder = lvwAscending Then
    PTY_LST.SortOrder = lvwDescending
ElseIf PTY_LST.SortOrder = lvwDescending Then
    PTY_LST.SortOrder = lvwAscending
End If
PTY_LST.SortKey = CH.Index - 1
End Sub

Private Sub PTY_LST_GotFocus()
PTY_LST.SetFocus
End Sub

Private Sub SMN_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Or Select Salesman"
End Sub

Private Sub smn_idy_kEYPRESS(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub SMN_LST_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If SMN_LST.SortOrder = lvwAscending Then
    SMN_LST.SortOrder = lvwDescending
ElseIf SMN_LST.SortOrder = lvwDescending Then
    SMN_LST.SortOrder = lvwAscending
End If
SMN_LST.SortKey = CH.Index - 1
End Sub

Private Sub smn_LST_DblClick()
Smn_LST_KeyPress 13
End Sub

Private Sub Smn_LST_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SMN_LST.ListItems.count > 0 Then
        SMN_IDY = SMN_LST.SelectedItem.ListSubItems(1)

        PTY_IDY = ""

        LDS_IDY = ""
        SMN_IDY.Enabled = False
        PTY_IDY.Enabled = False
        SMN_LST.Visible = False
        SSTab1.Tab = 0
        displaydetails
'        For A = 1 To MS.Rows - 1
'            If MS.TextMatrix(1, 0) = "" Then
'                MS.SetFocus
'            Else
'                MS.Col = 5
'            End If
'        Next
'        MS.SetFocus
    End If
End If
End Sub

Private Sub pty_LST_DblClick()
PTY_LST_KeyPress 13
End Sub

Private Sub PTY_LST_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If PTY_LST.ListItems.count > 0 Then
        PTY_IDY = PTY_LST.SelectedItem.ListSubItems(1)

        LDS_IDY = ""

        SMN_IDY = ""
        PTY_IDY.Enabled = False
        PTY_LST.Visible = False
        SSTab1.Tab = 0
        displaydetails
'        For A = 1 To MS.Rows - 1
'            If MS.TextMatrix(1, 0) = "" Then
'                MS.SetFocus
'            Else
'                MS.Col = 5
'            End If
'        Next
'        MS.SetFocus
    End If
Else
End If
End Sub

Private Sub saverecord()
Dim rs As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
Dim rec_vch As New ADODB.Recordset
Dim REC_SET_P As New ADODB.Recordset
If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
    Exit Sub
End If
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("ACT", "ACT_IDY", Trim(ACT_IDY)) = False Then
    MsgBox "Cash Account Not Found", vbExclamation, "DAS Version 4.0"
    ACT_IDY.SetFocus
    Exit Sub
End If
If check("ACT", "ACT_IDY", Trim(BNK_IDY)) = False Then
    MsgBox "Bank Account Not Found", vbExclamation, "DAS Version 4.0"
    BNK_IDY.SetFocus
    Exit Sub
End If
If DTPicker4.Visible = True Then
    MS.Text = DTPicker4.Value
    DTPicker4.Visible = False
End If
If Text1.Visible = True Then
    MS.Text = Trim(Text1)
    Text1.Visible = False
    Text1 = ""
End If
If PST_LST.Visible = True Then
    MS.Text = PST_LST.Text
    PST_LST.Visible = False
End If
If M_FLX_AMT.Visible = True Then
    MS.Text = Val(M_FLX_AMT.Text)
    M_FLX_AMT.Visible = False
    M_FLX_AMT = ""
End If
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 0)) <> "" Then
        con.Execute "UPDATE IRT SET CHK_FLG='N' WHERE INV_IDY LIKE '" & Trim(MS.TextMatrix(A, 1)) & "'", d
        If Trim(MS.TextMatrix(A, 5)) > 0 Then
            con.Execute "UPDATE IRT SET ACT_IDY='" & Trim(ACT_IDY) & "' WHERE INV_IDY LIKE '" & Trim(MS.TextMatrix(A, 1)) & "'"
        ElseIf Trim(MS.TextMatrix(A, 6)) > 0 Then
            con.Execute "UPDATE IRT SET ACT_IDY='" & Trim(BNK_IDY) & "' WHERE INV_IDY LIKE '" & Trim(MS.TextMatrix(A, 1)) & "'"
        End If
        con.Execute "UPDATE IRT SET ACT_BAL= " & Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 6)) & " WHERE INV_IDY = '" & Trim(MS.TextMatrix(A, 1)) & "'"
        con.Execute "UPDATE IRT SET RCD_AMT= " & (Val(MS.TextMatrix(A, 3)) + Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 6))) & " WHERE INV_IDY = '" & Trim(MS.TextMatrix(A, 1)) & "'"
        con.Execute "UPDATE IRT SET BNK_AMT= " & (Val(MS.TextMatrix(A, 6))) & " WHERE INV_IDY = '" & Trim(MS.TextMatrix(A, 1)) & "'"
    End If
Next
'NEXTRECORD
'RECEIPTS POSTING TO VCH
Counter = False
For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 5)) > 0 Then
        Counter = True
        Exit For
    End If
    Counter = False
Next
If Counter = True Then IDY = "RCT" & T7increment("IDY_RCT")
For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 5)) > 0 And Trim(MS.TextMatrix(A, 7)) = "Y" And Trim(MS.TextMatrix(A, 0)) <> "" Then
        If rec_vch.State = 1 Then rec_vch.Close
        rec_vch.Open "select * from VCH where txn_idy = ''", con, adOpenKeyset, adLockPessimistic
        rec_vch.AddNew
        rec_vch("TXN_idy") = IDY
        rec_vch("TXN_DAT") = DTPicker3
        rec_vch("TXN_AMT") = Val(MS.TextMatrix(A, 5))
        rec_vch("DBT_IDY") = Trim(ACT_IDY)
        rec_vch("CRT_IDY") = Trim(MS.TextMatrix(A, 11))
        rec_vch!CLG_DAT = (MS.TextMatrix(A, 12))
        rec_vch("TXN_RMK") = Trim(MS.TextMatrix(A, 13))
        rec_vch("TXN_TYP") = "RCT"
        rec_vch("INV_IDY") = Trim(MS.TextMatrix(A, 1))
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select TXN_TYP from IRT WHERE INV_IDY = '" & Trim(MS.TextMatrix(A, 1)) & "'", con, adOpenKeyset, adLockPessimistic
        If REC_SET.EOF = True Then
            rec_vch("INV_TYP") = "SAL"
        Else
            rec_vch("INV_TYP") = REC_SET("TXN_TYP")     'Checked and changed by PNR when IRT records do not match INL records due to inconsistent database (physical)
        End If
        rec_vch.Update
    End If
Next
Counter = False
For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 6)) > 0 Then
        Counter = True
        Exit For
    End If
    Counter = False
Next
If Counter = True Then IDY = "RCT" & T7increment("IDY_RCT")
For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 6)) > 0 And Trim(MS.TextMatrix(A, 7)) = "Y" And Trim(MS.TextMatrix(A, 0)) <> "" Then
        If rec_vch.State = 1 Then rec_vch.Close
        rec_vch.Open "select * from VCH where txn_idy = ''", con, adOpenKeyset, adLockPessimistic
        rec_vch.AddNew
        rec_vch("TXN_idy") = IDY
        rec_vch("TXN_DAT") = DTPicker3
        rec_vch("TXN_AMT") = Val(MS.TextMatrix(A, 6))
        rec_vch("DBT_IDY") = Trim(BNK_IDY)
        rec_vch("CRT_IDY") = Trim(MS.TextMatrix(A, 11))
        rec_vch!CLG_DAT = (MS.TextMatrix(A, 12))
        rec_vch("TXN_RMK") = Trim(MS.TextMatrix(A, 13))
        rec_vch("TXN_TYP") = "RCT"
        rec_vch("INV_IDY") = Trim(MS.TextMatrix(A, 1))
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select TXN_TYP from IRT WHERE INV_IDY = '" & Trim(MS.TextMatrix(A, 1)) & "'", con, adOpenKeyset, adLockPessimistic
        If REC_SET.EOF = True Then
            rec_vch("INV_TYP") = "SAL"
        Else
            rec_vch("INV_TYP") = REC_SET("TXN_TYP") 'Checked and changed by PNR when IRT records do not match INL records due to inconsistent database (physical)
        End If
        rec_vch.Update
    End If
Next
If CHECKTB("inl.dbf,prt.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If Val(REC_LBL) > 0 Then MsgBox "Account Receipts Posted by System! Do not Post Again from Receipts!", vbExclamation, "DAS Version 4.0"
SAL_QRY = "INL->TXN_TYP = 'SAL' and IRT->CHK_FLG = 'N'"
If check("PRT", "PRT_NME", "M_FRM_IRT") = True Then
    'If smn_chk.Value = 1 Then
     If Combo1.ListIndex = 1 Then 'Added by PNR
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy = 'SMN' AND dir_idx = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
                rs("DIR_ALS") = Trim(SMN_IDY) & ""
                rs.Update
        End If
        If Len(Trim(SAL_QRY)) = 0 Then
            SAL_QRY = "IRT->SMN_IDY='" & Trim(SMN_IDY) & "'"
        Else
            SAL_QRY = SAL_QRY & ".AND.IRT->SMN_IDY='" & Trim(SMN_IDY) & "'"
        End If
    Else
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy = 'SMN' AND dir_idx = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
                rs("DIR_ALS") = "*"
                rs("DIR_MSG") = "All"
                rs.Update
        End If
    End If
    
    'If PTY_CHK.Value = 1 Then
     If Combo1.ListIndex = 2 Then 'Added by PNR
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy = 'DLR' AND dir_idx = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
                rs("DIR_ALS") = Trim(PTY_IDY) & ""
                rs.Update
        End If
        If Len(Trim(SAL_QRY)) = 0 Then
            SAL_QRY = "INL->CST_IDY='" & Trim(PTY_IDY) & "'"
        Else
            SAL_QRY = SAL_QRY & ".AND.INL->CST_IDY='" & Trim(PTY_IDY) & "'"
        End If
    Else
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy = 'DLR' AND dir_idx = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
                rs("DIR_ALS") = "*"
                rs("DIR_MSG") = "All"
                rs.Update
        End If
    End If
    
    'If LDS_CHK.Value = 1 Then
    If Combo1.ListIndex = 3 Then 'Added by PNR
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy = 'LOC' AND dir_idx = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
                rs("DIR_ALS") = Trim(LDS_IDY) & ""
                rs.Update
        End If
    Else
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy = 'LOC' AND dir_idx = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
                rs("DIR_ALS") = "*"
                rs("DIR_MSG") = "All"
                rs.Update
        End If
    End If
    'If VAN_CHK.Value = 1 Then
    If Combo1.ListIndex = 4 Then 'Added by PNR
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy = 'CAT' AND dir_idx = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
                rs("DIR_ALS") = Trim(VAN_IDY) & ""
                rs.Update
        End If
    Else
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy = 'CAT' AND dir_idx = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
                rs("DIR_ALS") = "*"
                rs("DIR_MSG") = "All"
                rs.Update
        End If
    End If
    
    If Combo1.ListIndex = 5 Then 'Added by PNR
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy = 'INV' AND dir_idx = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            rs("DIR_ALS") = Trim(M_INV_IDY) & ""
            rs.Update
        End If
    Else
        If rs.State = 1 Then rs.Close
        rs.Open "select * from dir where dir_idy = 'INV' AND dir_idx = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
                rs("DIR_ALS") = "*"
                rs("DIR_MSG") = "All"
                rs.Update
        End If
    End If
        If Len(Trim(SAL_QRY)) = 0 Then
            'SAL_QRY = "(IRT->NET_AMT - IRT->RCD_AMT) <> 0 .and. (ACT_BAL >= 0)"
        Else
            'SAL_QRY = SAL_QRY & ".AND. (IRT->NET_AMT - IRT->RCD_AMT) <> 0 .AND. (ACT_BAL >= 0)"
        End If
    OpenReport CURDIR & "\REP_LIB.RP1", "Day Account Sheet Check", 2, SAL_QRY, 6
End If
'END OF POSTING
clrctr M_FRM_IRT
setval
CANCELRECORD
SSTab1.Tab = 0
DTPicker3.SetFocus
MOD_REC = False
End Sub

Private Sub CANCELRECORD()
MOD_REC = False
Command1.Enabled = True
setval
Form_Load
MS.Rows = 1
SSTab1.Tab = 0
DTPicker1.Enabled = True
DTPicker2.Enabled = True
DTPicker1.SetFocus
ALL_CHK.Enabled = True
ALL_CHK.Value = 0 'Changed by PNR to 0 from 1 BASED on Prince Chopra Feedback
ACT_LBL = ""
BNK_LBL = ""
asdf = "Receipt Date"
PST_LST.Visible = False
DTPicker4.Visible = False
Text1.Visible = False
M_FLX_AMT.Visible = False
SMN_IDY.Enabled = True
PTY_IDY.Enabled = True
LDS_IDY.Enabled = True
VAN_IDY.Enabled = True
DTPicker1.Enabled = True
DTPicker2.Enabled = True
Command1.Enabled = False
Combo1.Enabled = True
End Sub

Public Sub setval()
clrctr M_FRM_IRT
ALL_CHK.Enabled = True
ALL_CHK.Value = 0 'Changed by PNR based on Prince Chopra feedback
LDS_IDY = ""
VAN_IDY = ""
SMN_IDY = ""
PTY_IDY = ""
REC_LBL.Caption = "0.00"
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub

Private Sub Smn_IDY_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If KeyCode = 13 Then
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
    rs.Open "select * from smn where smn_IDY = '" & Sincrement(Trim(SMN_IDY)) & "'", con, adOpenStatic
    If rs.RecordCount > 0 Then
        SMN_IDY = Sincrement(Trim(SMN_IDY))
        displaydetails

        PTY_IDY = ""

        VAN_IDY = ""

        LDS_IDY = ""
    Else
        V_SMN_LST
    End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    newrec
ElseIf Button.KEY = "s" Then
    saverecord
ElseIf Button.KEY = "c" Then
    CANCELRECORD
ElseIf Button.KEY = "e" Then
    Unload Me
    Set Form = Nothing
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "p" Then
    Printrecord
ElseIf Button.KEY = "idc" Then
    If Button.Value = tbrPressed Then
        Option1.Value = True
    Else
        Option1.Value = False
    End If
ElseIf Button.KEY = "ed" Then
    If M_FRM_IRT.ActiveControl.name = "ACT_IDY" Then V_ACT_LST
    If M_FRM_IRT.ActiveControl.name = "BNK_IDY" Then V_bnk_LST
    If M_FRM_IRT.ActiveControl.name = "SMN_IDY" Then V_SMN_LST
    If M_FRM_IRT.ActiveControl.name = "PTY_IDY" Then V_PTY_LST
    If M_FRM_IRT.ActiveControl.name = "LDS_IDY" Then V_LDS_LST
    If M_FRM_IRT.ActiveControl.name = "VAN_IDY" Then V_RSI_LST
ElseIf Button.KEY = "pr" Then
    Exit Sub
er1:
    MsgBox "Account Not Found, Define Account...!", vbExclamation, "DAS Version 4.0"
End If
End Sub

Public Sub newrec()
CANCELRECORD
PTY_IDY.Enabled = True
SMN_IDY.Enabled = True
DTPicker1.SetFocus
SSTab1.Tab = 0
MOD_REC = False
End Sub

Public Sub Printrecord()
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_IRT") = True Then
    M_FRM_PRV.Check3.Value = 1
Else
    M_FRM_PRV.Check3.Value = 0
End If
M_FRM_PRV.Tag = "IRT"
M_FRM_PRV.Show 1
End Sub

Public Function CHKINL(INVIDY) As Boolean
For A = 1 To MS.Rows - 1
    If Trim(MS.TextMatrix(A, 1)) = INVIDY Then
        CHKINL = True
        Exit Function
    End If
Next
CHKINL = False
End Function
Private Sub VAN_LST_DblClick()
VAN_LST_KeyPress 13
End Sub
Private Sub VAN_LST_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    VAN_IDY = VAN_LST.SelectedItem
    VAN_LST.Visible = False
    
    PTY_IDY = ""
    LDS_IDY = ""
    
    
    ALL_CHK.Value = 0
    SMN_IDY = ""
    SSTab1.Tab = 0
    VAN_IDY.Enabled = False
    displaydetails
End If
End Sub
Public Sub Total()
'If MS.Col = 5 Or MS.Col = 6 Then
    REC_LBL = "0.00"
    Bnk_amt = "0.00"
    Bal_Amt = "0.00"
    For A = 1 To MS.Rows - 1
        REC_LBL = convert(Val(REC_LBL) + Val(MS.TextMatrix(A, 5)))
        Bnk_amt = convert(Val(Bnk_amt) + Val(MS.TextMatrix(A, 6)))
        If Val(MS.TextMatrix(A, 4)) >= 0 Then
            Bal_Amt = convert(Val(Bal_Amt) + Val(MS.TextMatrix(A, 4)))
        End If
    Next
'End If
End Sub
