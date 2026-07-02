VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_QPS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quantitative Purchase Scheme"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "m_frm_QPS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8760
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
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "&QPS"
      TabPicture(0)   =   "m_frm_QPS.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "retname"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MS"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "retid"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtvar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "&Details"
      TabPicture(1)   =   "m_frm_QPS.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TXTSAL"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MST"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "sbalamt"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "spaidamt"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ssdisamt"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ssqty"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "sdisamt"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "stqty"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "&Payments"
      TabPicture(2)   =   "m_frm_QPS.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Qpsamt"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label7"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "&View"
      TabPicture(3)   =   "m_frm_QPS.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label20"
      Tab(3).Control(1)=   "lv"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   90
         TabIndex        =   61
         Top             =   1440
         Width           =   8630
         Begin VB.CheckBox All 
            Caption         =   "&All Products"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   315
            Width           =   1215
         End
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            MaxLength       =   8
            TabIndex        =   63
            TabStop         =   0   'False
            Text            =   " "
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txtDisamt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   5880
            MaxLength       =   11
            TabIndex        =   62
            TabStop         =   0   'False
            Text            =   " "
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Quantity"
            Height          =   195
            Left            =   1920
            TabIndex        =   66
            Top             =   345
            Width           =   585
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Discount Amt Per Case"
            Height          =   195
            Left            =   4080
            TabIndex        =   65
            Top             =   345
            Width           =   1635
         End
      End
      Begin VB.Frame Frame1 
         Height          =   700
         Left            =   5040
         TabIndex        =   56
         Top             =   720
         Width           =   3670
         Begin MSComCtl2.DTPicker fdate 
            Height          =   330
            Left            =   540
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   235
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37869
         End
         Begin MSComCtl2.DTPicker tdate 
            Height          =   330
            Left            =   2190
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   235
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37869
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "To"
            Height          =   195
            Left            =   1900
            TabIndex        =   60
            Top             =   303
            Width           =   195
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "From"
            Height          =   195
            Left            =   110
            TabIndex        =   58
            Top             =   303
            Width           =   345
         End
      End
      Begin VB.TextBox txtvar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Qpsamt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72890
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   430
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Payment Details"
         Height          =   4335
         Left            =   -74640
         TabIndex        =   15
         Top             =   720
         Width           =   5415
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   11
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "0.00"
            Top             =   3960
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   10
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   53
            Text            =   "0.00"
            Top             =   3620
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   9
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "0.00"
            Top             =   3282
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   8
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "0.00"
            Top             =   2944
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "0.00"
            Top             =   2606
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "0.00"
            Top             =   2268
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "0.00"
            Top             =   1930
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   47
            Text            =   "0.00"
            Top             =   1592
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "0.00"
            Top             =   1254
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "0.00"
            Top             =   916
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "0.00"
            Top             =   578
            Width           =   1695
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "0.00"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   11
            Left            =   1750
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   3960
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   10
            Left            =   1750
            TabIndex        =   41
            Text            =   "0.00"
            Top             =   3620
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   9
            Left            =   1750
            TabIndex        =   40
            Text            =   "0.00"
            Top             =   3282
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   8
            Left            =   1750
            TabIndex        =   39
            Text            =   "0.00"
            Top             =   2944
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   1750
            TabIndex        =   38
            Text            =   "0.00"
            Top             =   2606
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   1750
            TabIndex        =   37
            Text            =   "0.00"
            Top             =   2268
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   1750
            TabIndex        =   36
            Text            =   "0.00"
            Top             =   1930
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1750
            TabIndex        =   35
            Text            =   "0.00"
            Top             =   1592
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1750
            TabIndex        =   34
            Text            =   "0.00"
            Top             =   1254
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1750
            TabIndex        =   33
            Text            =   "0.00"
            Top             =   916
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1750
            TabIndex        =   32
            Text            =   "0.00"
            Top             =   578
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1750
            TabIndex        =   31
            Text            =   "0.00"
            Top             =   240
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   275
            Index           =   9
            Left            =   120
            TabIndex        =   18
            Top             =   3282
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   275
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   916
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   275
            Index           =   8
            Left            =   120
            TabIndex        =   20
            Top             =   2944
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   275
            Index           =   7
            Left            =   120
            TabIndex        =   21
            Top             =   2606
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   275
            Index           =   6
            Left            =   120
            TabIndex        =   22
            Top             =   2268
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   275
            Index           =   5
            Left            =   120
            TabIndex        =   23
            Top             =   1930
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   275
            Index           =   4
            Left            =   120
            TabIndex        =   24
            Top             =   1592
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   275
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   1254
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   275
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   578
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   275
            Index           =   10
            Left            =   120
            TabIndex        =   28
            Top             =   3620
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
         Begin MSComCtl2.DTPicker pdate 
            Height          =   275
            Index           =   11
            Left            =   120
            TabIndex        =   29
            Top             =   3960
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37879
         End
      End
      Begin VB.TextBox TXTSAL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73920
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox retid 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MST 
         Height          =   4300
         Left            =   -74940
         TabIndex        =   5
         Top             =   480
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   7594
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
      Begin MSComctlLib.ListView lv 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1155
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
         Left            =   50
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   2280
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   7
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "QPS/UPS Amount"
         Height          =   195
         Left            =   -74520
         TabIndex        =   16
         Top             =   475
         Width           =   1320
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
         Left            =   -74550
         TabIndex        =   14
         Top             =   800
         Width           =   7905
      End
      Begin VB.Label sbalamt 
         Caption         =   " "
         Height          =   255
         Left            =   -67080
         TabIndex        =   12
         Top             =   4845
         Width           =   855
      End
      Begin VB.Label spaidamt 
         Caption         =   " "
         Height          =   255
         Left            =   -67920
         TabIndex        =   11
         Top             =   4845
         Width           =   855
      End
      Begin VB.Label ssdisamt 
         Caption         =   " "
         Height          =   255
         Left            =   -68520
         TabIndex        =   10
         Top             =   4845
         Width           =   855
      End
      Begin VB.Label ssqty 
         Caption         =   " "
         Height          =   255
         Left            =   -69960
         TabIndex        =   9
         Top             =   4845
         Width           =   855
      End
      Begin VB.Label sdisamt 
         Caption         =   " "
         Height          =   255
         Left            =   -71160
         TabIndex        =   8
         Top             =   4845
         Width           =   855
      End
      Begin VB.Label stqty 
         Height          =   255
         Left            =   -72180
         TabIndex        =   7
         Top             =   4845
         Width           =   975
      End
      Begin VB.Label retname 
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
         Left            =   2520
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Retailer"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   1155
         Width           =   1035
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
            Picture         =   "m_frm_QPS.frx":04B2
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":09F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":0DBE
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":1112
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":174A
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":1A9E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":1EFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":21F2
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":25FE
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":270A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":2A5E
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":2E72
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
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
            Picture         =   "m_frm_QPS.frx":33BE
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":3902
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":3CCA
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":401E
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":4656
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":49AA
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":4E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":50FE
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":550A
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":5616
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":596A
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "m_frm_QPS.frx":5D7E
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "m_frm_qps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim rsact As New ADODB.Recordset
Dim rsqps As New ADODB.Recordset
Dim RSPRD As New ADODB.Recordset
Dim RSMAR As New ADODB.Recordset
Dim RSINV As New ADODB.Recordset
Dim MD As Boolean, R As Integer, S As Integer
Private Sub All_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If All.Value = 1 Then
        tog (True)
        txtQty.SetFocus
        SendKeys "{home}+{end}"
    Else
        tog (False)
        dispprd
        MS.SetFocus
        MS.Col = 1
    End If
End If
End Sub

Private Sub Form_Load()
SETYEAR Me
SSTAB1.TabEnabled(3) = False
DISPPH
End Sub

Private Sub MS_Scroll()
If txtvar.Visible = True Then
    txtvar = ""
    txtvar.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub TEXT2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Index = 0 Then
        'bal1(Index).Locked = True
        Text2(Index) = convert(Text2(Index))
        Text3(Index) = convert(Val(Qpsamt.Text) - Val(Text2(Index)))
        'bal1(Index).Enabled = True
        pdate(Index + 1).Enabled = True
        pdate(Index + 1).SetFocus
    Else
        'bal1(Index).Locked = True
        Text2(Index) = convert(Text2(Index))
        Text3(Index) = convert(Val(Text3(Index - 1).Text) - Val(Text2(Index)))
        'bal1(Index).Enabled = True
        pdate(Index + 1).Enabled = True
        pdate(Index + 1).SetFocus
        SendKeys "{HOME}+{END}"
    End If
End If
End Sub

Private Sub bal1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Index < 11 Then
        'pdate(Index + 1).Enabled = True
        pdate(Index + 1).SetFocus
    End If
End If
End Sub

Private Sub FDATE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    tdate.SetFocus
End If
End Sub

Private Sub Form_Activate()
'txnid.SetFocus
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
    '    qpslist
    'Else
    If Me.ActiveControl.name = "retid" Then
        retlist
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    'If Me.ActiveControl.name = "txnid" And Len(Trim(txnid)) = 0 Then
     '   Unload Me
    'Else
    If SSTAB1.Tab = 3 Then
        If lv.Visible = True And UCase(Trim(Label20)) = UCase("Select Transaction") Then
            lv.Visible = False
            SSTAB1.TabEnabled(3) = False
            SSTAB1.Tab = 0
            'If txnid.Enabled = True Then
            '    txnid.SetFocus
            'Else
            '    retid.SetFocus
            'End If
        ElseIf lv.Visible = True And UCase(Trim(Label20)) = UCase("Select Retailer") Then
            lv.Visible = False
            SSTAB1.Tab = 2
            Qpsamt.SetFocus
        End If
    ElseIf SSTAB1.Tab = 1 Then
        SSTAB1.Tab = 0
        'If txnid.Enabled = True Then
        '    txnid.SetFocus
        'Else
        If retid.Enabled = True Then
            retid.SetFocus
        End If
    ElseIf SSTAB1.Tab = 2 Then
        SSTAB1.Tab = 1
        MS.SetFocus
    Else
        saverecord
    End If
End If
End Sub

Private Sub lv_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If lv.SortOrder = lvwAscending Then
lv.SortOrder = lvwDescending
ElseIf lv.SortOrder = lvwDescending Then
lv.SortOrder = lvwAscending
End If
lv.SortKey = CH.Index - 1
End Sub

Private Sub lv_DblClick()
lv_KeyDown 13, 0
End Sub

Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RS As New ADODB.Recordset
If KeyCode = 13 Then
    If lv.ListItems.Count > 0 Then
        If RS.State = 1 Then RS.Close
        If Trim(Label20.Caption) = "Select Retailer" Then
            Dim lst As ListSubItem
            Set lst = lv.SelectedItem.ListSubItems.Item(1)
            RS.Open "select * from act where act_idy like '" & Trim(lst.Text) & "'", CON, adOpenStatic
            retid = Trim(RS("act_idy"))
            retname = Trim(RS("act_nme"))
            lv.Visible = False
            SSTAB1.TabEnabled(3) = False
            SSTAB1.Tab = 0
            fdate.SetFocus
            retid.Enabled = False
        Else
            Dim ls As ListItem
            Set ls = lv.SelectedItem
            RS.Open "select * from qps where txn_idy like '" & Trim(ls.Text) & "'", CON, adOpenKeyset, adLockOptimistic
            If RS.RecordCount > 0 Then
'            txnid = Sincrement(txnid)
                If RS("CHK_FLG") = "A" Or RS("CHK_FLG") = "S" Then
                    'rs.Open "select * from inc where txn_idy like '" & Trim(txnid) & "'", CON, adOpenStatic
                    If rsact.State = 1 Then rsact.Close
                    rsact.Open "select act_nme from act where act_idy ='" & Trim(RS("act_idy")) & "'", CON, adOpenKeyset, adLockOptimistic
                    If Not rsact.EOF Then
                        retid = RS("act_idy")
                        retname.Caption = rsact(0)
                    End If
                    fdate.Value = RS("frm_dat")
                    tdate.Value = RS("to_dat")
                    All.Enabled = False
                    DISPT
                    If RS("CHK_FLG") = "A" Then
                    R = 1
                        All.Value = 1
                        MST.Rows = MST.Rows + 1
                        MST.Cols = 9
                        txtQty = RS("PRD_QTY")
                        txtDisamt = convert(RS("dis_AMT"))
                        MST.TextMatrix(R, 0) = ""
                        MST.TextMatrix(R, 1) = RS("PRD_QTY")
                        MST.TextMatrix(R, 2) = Val(txtDisamt)
                        If RSINV.State = 1 Then RSINV.Close
                        RSINV.Open "select sum(tot_pcs) as tot_pcs from inv where inv_idy in(select inv_idy from inl where inv_dat>=ctod('" & fdate.Value & "') and inv_dat<=ctod('" & tdate.Value & "') and cst_idy='" & Trim(retid) & "') AND stk_typ in('Salable','Free')", CON, adOpenKeyset, adLockOptimistic
                        If Not RSINV.EOF Then
                            MST.TextMatrix(R, 3) = Val(RSINV(0))
                            If Val(MST.TextMatrix(R, 1)) <= Val(MST.TextMatrix(R, 3)) Then
                                MST.TextMatrix(R, 4) = Val(Trim(MST.TextMatrix(R, 3))) * Val(txtDisamt)
                            End If
                        End If
'                        txnid.Enabled = False
                        For d = 0 To 11
                            If Val(RS(d + 20)) >= 0 Then
                                If d = 0 Then
                                    pdate(d).Enabled = True
                                    pdate(d).Value = RS(d + 8)
                                    Text2(d) = convert(RS(d + 20))
                                Else
                                    pdate(d).Enabled = True
                                    pdate(d).Value = RS(d + 8)
                                    Text2(d).Text = convert(RS(d + 20))
                                End If
                            End If
                        Next
                    Else
                        All.Value = 0
                        R = 1
                        MS.Rows = 1
                        MS.Cols = 5
                        R = 1
                        MS.ColWidth(3) = 0
                        MS.ColWidth(4) = 0
                        MS.ColWidth(0) = 2500
                        MS.ColWidth(1) = 1100
                        MS.ColWidth(2) = 1800
                        MS.TextMatrix(0, 0) = "Product Name"
                        MS.TextMatrix(0, 1) = "Quantity"
                        MS.ColAlignment(1) = 7
                        SCl MS, 1
                        MS.TextMatrix(0, 2) = "Discount Amt Per Case"
                        SCl MS, 2
                        MS.ColAlignment(2) = 7
                        If RSINV.State = 1 Then RSINV.Close
                        RSINV.Open "select sum(tot_pcs) as tot_pcs,sub_unt,prd_idy from inv where inv_idy in(select inv_idy from inl where inv_dat>=ctod('" & fdate.Value & "') and inv_dat<=ctod('" & tdate.Value & "') and cst_idy='" & Trim(retid) & "') AND stk_typ in('Salable','Free') group by prd_idy", CON, adOpenKeyset, adLockOptimistic
                        While Not RS.EOF
                            If Not RSINV.EOF Then RSINV.MoveFirst
                            MS.Rows = MS.Rows + 1
                            MST.Rows = MST.Rows + 1
                            If RSPRD.State = 1 Then RSPRD.Close
                            RSPRD.Open "select prd_nme,SUB_UNT from prd where prd_idy='" & RS("prd_idy") & "'", CON, adOpenKeyset, adLockOptimistic
                            MS.TextMatrix(R, 0) = RSPRD(0)
                            MST.TextMatrix(R, 0) = RSPRD(0)
                            MS.TextMatrix(R, 4) = RSPRD(1)
                            MST.TextMatrix(R, 6) = RSPRD(1)
                            MS.TextMatrix(R, 3) = RS("prd_idy")
                            MST.TextMatrix(R, 7) = RS("prd_idy")
                            MS.TextMatrix(R, 1) = QTY(RS("prd_qty"))
                            MST.TextMatrix(R, 1) = QTY(RS("prd_qty"))
                            MS.TextMatrix(R, 2) = convert(RS("dis_amt"))
                            MST.TextMatrix(R, 2) = convert(RS("DIS_AMT"))
                            Do While Not RSINV.EOF
                                If RS("prd_idy") = RSINV("prd_idy") Then
                                    MST.TextMatrix(R, 3) = STM(Val(RSINV(0)), Val(RSINV(1)))
                                    'MST.TextMatrix(r, 4) = convert(CASEAMT(Trim(MS.TextMatrix(r, 3)), Trim(MS.TextMatrix(r, 4)), Trim(MS.TextMatrix(r, 2))))
                                    Exit Do
                                End If
                                RSINV.MoveNext
                            Loop
                            R = R + 1
                            RS.MoveNext
                        Wend
                        If RS.RecordCount > 0 Then
                        RS.MoveFirst
                        For d = 0 To 11
                            V = d
                            If Val(RS(d + 20)) > 0 Then
                                If d = 0 Then
                                    pdate(d).Enabled = True
                                    pdate(d).Value = RS(d + 8)
                                    Text2(d) = convert(RS(d + 20))
                                    'Text3(d) = Val(Qpsamt.Text) - Val(Text2(d))
                                Else
                                    pdate(d).Enabled = True
                                    pdate(d).Value = RS(d + 8)
                                    Text2(d) = convert(RS(d + 20))
                                    'Text3(d) = Val(Text3(d - 1)) - Val(Text2(d))
                                End If
                            Else
                                Exit For
                            End If
                        Next
                        End If
                        For R = 1 To MST.Rows - 1
                            If Val(MST.TextMatrix(R, 1)) >= Val(MST.TextMatrix(R, 3)) Then
                                Exit For
                            End If
                        Next
                        If R = MST.Rows Then
                            For R = 1 To MST.Rows - 1
                                MST.TextMatrix(R, 4) = convert(CASEAMT(Trim(MST.TextMatrix(R, 3)), Trim(MS.TextMatrix(R, 6)), Trim(MS.TextMatrix(R, 2))))
                            Next
                        End If
                            'smnid.Enabled = False
                        MS.SetFocus
                        MS.Col = 2
                    End If
                    lv.Visible = False
                    SSTAB1.Tab = 0
                    SSTAB1.TabEnabled(3) = False
                    retid.SetFocus
                    SendKeys "{HOME}+{END}"
                    MD = True
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub MS_KeyPress(KeyAscii As Integer)
If keyasii = 13 Or KeyAscii = 27 Then

ElseIf KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Then
     If MS.Col = 1 Or MS.Col = 2 Then
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

Private Sub MST_KeyPress(KeyAscii As Integer)
'If KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 46 Then
'    If MST.Col = 3 Then
'        TXTSAL.Visible = True
'        TXTSAL.Left = MST.Left + MST.CellLeft
'        TXTSAL.Top = MST.Top + MST.CellTop
'        TXTSAL.Height = MST.CellHeight
'        TXTSAL.Width = MST.CellWidth
'        TXTSAL.SetFocus
'        TXTSAL = Chr(KeyAscii)
'        TXTSAL.SelStart = Len(Trim(TXTSAL.Text))
'    End If
'End If
End Sub

Private Sub pdate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
pdate(Index).Refresh
If KeyCode = 13 Then
    If Index > 0 Then
        If pdate(Index).Value >= pdate(Index - 1).Value Then
            Text2(Index).SetFocus
            If Val(Text2(Index)) <= 0 Then Text2(Index) = "0.00"
            SendKeys "{HOME}+{END}"
        Else
            MsgBox "Invalid Date", vbExclamation
            pdate(Index).SetFocus
        End If
    Else
        Text2(Index).SetFocus
        If Val(Text2(Index)) <= 0 Then Text2(Index) = "0.00"
        SendKeys "{HOME}+{END}"
    End If
End If
End Sub

Private Sub Qpsamt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Val(Qpsamt) > 0 Then
        pdate(0).Enabled = True
        pdate(0).SetFocus
        For d = 0 To 11
            If d = 0 Then
                If Val(Text2(d)) >= 0 Then
                    Text3(d) = convert(Val(Qpsamt.Text) - Val(Text2(d)))
                End If
            Else
                If Val(Text2(d)) > 0 Then
                    Text3(d) = convert(Val(Text3(d - 1)) - Val(Text2(d)))
                End If
            End If
        Next
    End If
End If
End Sub

Private Sub retid_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim RS As New ADODB.Recordset
If K = 13 Or K = 40 Then
    If RS.State = 1 Then RS.Close
    RS.Open "select * from act where act_idy like '" & Sincrement(retid) & "' AND grp_clf='Debtor'", CON, adOpenStatic
    If RS.RecordCount > 0 Then
        retid = Sincrement(retid)
        retname = Trim(RS("act_nme"))
        retid.Enabled = False
        fdate.SetFocus
    Else
        retlist
    End If
End If
End Sub
Private Sub retlist()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
Me.MousePointer = vbHourglass
RS.Open "select * from act where act_NME like '" & retid & "%' AND grp_clf='Debtor'", CON, adOpenStatic
Label20.Caption = "Select Retailer"
lv.ColumnHeaders.Clear
lv.ColumnHeaders.Add , , "Retailer Name", 2500
lv.ColumnHeaders.Add , , "ID", 1500
lv.ColumnHeaders.Add , , "Market Name", 2500
lv.ListItems.Clear
If RS.RecordCount > 0 Then
    Me.MousePointer = vbHourglass
    For A = 1 To RS.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(RS("act_nme")))
        ls.ListSubItems.Add , , Trim(RS("act_idy") & "")
        If RSMAR.State = 1 Then RSMAR.Close
        RSMAR.Open "SELECT MAR_NME FROM MAR WHERE MAR_IDY='" & RS("MAR_IDY") & "'", CON, adOpenKeyset, adLockBatchOptimistic
        If Not RSMAR.EOF Then
            ls.ListSubItems.Add , , Trim(RSMAR("MAR_NME")) & ""
        End If
        RS.MoveNext
    Next
    SSTAB1.TabEnabled(3) = True
    SSTAB1.Tab = 3
    lv.Visible = True
    lv.SetFocus
    Me.MousePointer = vbNormal
Else
    If SSTAB1.Tab = 0 Then
        MsgBox "Retailer Not Found, Define Retailer!", vbExclamation, "Effmcg"
        SSTAB1.Tab = 0
        retid.Enabled = True
        retid.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End If
Me.MousePointer = vbNormal
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
stqty = "0.000"
ssqty = "0.000"
ssdisamt = "0.00"
If SSTAB1.Tab = 1 Then
    'If Len(Trim(txnid)) = 0 Then
    '    DISPT
    'Else
    MST.SetFocus
    MST.Col = 4
    For R = 1 To MST.Rows - 1
        stqty = QTY(Val(stqty) + Val(MST.TextMatrix(R, 1)))
        'sdisamt = convert(Val(sdisamt) + Val(MST.TextMatrix(r, 2)))
        ssqty = QTY(Val(ssqty) + Val(MST.TextMatrix(R, 3)))
        ssdisamt = convert(Val(ssdisamt) + Val(MST.TextMatrix(R, 4)))
'        spaidamt = convert(Val(spaidamt) + Val(MST.TextMatrix(r, 5)))
 '       sbalamt = convert(Val(sbalamt) + Val(MST.TextMatrix(r, 8)))
    Next
    'End If
ElseIf SSTAB1.Tab = 2 Then
'    If Len(Trim(txnid)) = 0 Then
'        DISPT
'    Else
'
    If MST.Rows > 1 Then MST.SetFocus
        MST.Col = 4
    For R = 1 To MST.Rows - 1
        stqty = QTY(Val(stqty) + Val(MST.TextMatrix(R, 1)))
        'sdisamt = convert(Val(sdisamt) + Val(MST.TextMatrix(r, 2)))
        ssqty = QTY(Val(ssqty) + Val(MST.TextMatrix(R, 3)))
        ssdisamt = convert(Val(ssdisamt) + Val(MST.TextMatrix(R, 4)))
'        spaidamt = convert(Val(spaidamt) + Val(MST.TextMatrix(r, 5)))
 '       sbalamt = convert(Val(sbalamt) + Val(MST.TextMatrix(r, 8)))
    Next
    Qpsamt.Text = convert(Trim(ssdisamt.Caption))
    Qpsamt.SetFocus
    SendKeys "{HOME}+{END}"
    'End If
End If
End Sub

Private Sub TDATE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If fdate.Value > tdate.Value Then
        MsgBox "Invalid Date ", vbExclamation
        fdate.SetFocus
        Exit Sub
    End If
    If All.Enabled = True Then
        All.SetFocus
    ElseIf All.Value = 0 Then
        MS.SetFocus
        MS.Col = 1
    Else
        txtQty.Enabled = True
        txtQty.SetFocus
        SendKeys "{HOME}+{END}"
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
  '  MsgBox Me.ActiveControl.name
    If Me.ActiveControl.name = "retid" Then
        retlist
'    ElseIf Me.ActiveControl.name = "txnid" Then
'        qpslist
    End If
End If
End Sub

'Private Sub txnid_KeyDown(K As Integer, Shift As Integer)
'If CHECKTB("qps.dbf") = 1 Then
'    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'    Exit Sub
'End If
'Dim RS As New ADODB.Recordset
'If K = 13 Or K = 40 Then
'    If RS.State = 1 Then RS.Close
'    RS.Open "select * from qps where txn_idy like '" & Sincrement(txnid) & "'", CON, adOpenStatic
'    If RS.RecordCount > 0 Then
'        If RS("CHK_FLG") = "A" Or RS("CHK_FLG") = "S" Then
'            txnid = Sincrement(txnid)
'            'rs.Open "select * from inc where txn_idy like '" & Trim(txnid) & "'", CON, adOpenStatic
'            If rsact.State = 1 Then rsact.Close
'            rsact.Open "select act_nme from act where act_idy ='" & Trim(RS("act_idy")) & "'", CON, adOpenKeyset, adLockOptimistic
'            If Not rsact.EOF Then
'                retid = RS("act_idy")
'                retname.Caption = rsact(0)
'            End If
'            fdate.Value = RS("frm_dat")
'            tdate.Value = RS("to_dat")
'            All.Enabled = False
'            DISPT
'            If RS("CHK_FLG") = "A" Then
'            R = 1
'                All.Value = 1
'                MST.Rows = MST.Rows + 1
'                txtQty = RS("PRD_QTY")
'                txtDisamt = convert(RS("dis_AMT"))
'                MST.TextMatrix(R, 0) = ""
'                MST.TextMatrix(R, 1) = RS("PRD_QTY")
'                MST.TextMatrix(R, 2) = Val(txtDisamt)
'                If RSINV.State = 1 Then RSINV.Close
'                RSINV.Open "select sum(tot_pcs) as tot_pcs from inv where inv_idy in(select inv_idy from inl where inv_dat>=ctod('" & fdate.Value & "') and inv_dat<=ctod('" & tdate.Value & "') and cst_idy='" & Trim(retid) & "') AND stk_typ in('Salable','Free')", CON, adOpenKeyset, adLockOptimistic
'                If Not RSINV.EOF Then
'                    MST.TextMatrix(R, 3) = Val(RSINV(0))
'                    If Val(MST.TextMatrix(R, 1)) <= Val(MST.TextMatrix(R, 3)) Then
'                        MST.TextMatrix(R, 4) = Val(Trim(MST.TextMatrix(R, 3))) * Val(txtDisamt)
'                    End If
'                End If
'                txnid.Enabled = False
'                For d = 0 To 11
'                    If Val(RS(d + 20)) >= 0 Then
'                        If d = 0 Then
'                            pdate(d).Enabled = True
'                            pdate(d).Value = RS(d + 8)
'                            Text2(d) = convert(RS(d + 20))
'                        Else
'                            pdate(d).Enabled = True
'                            pdate(d).Value = RS(d + 8)
'                            Text2(d).Text = convert(RS(d + 20))
'                        End If
'                    End If
'                Next
'            Else
'                All.Value = 0
'                R = 1
'                MS.Rows = 1
'                MS.Cols = 5
'                MS.ColWidth(3) = 0
'                MS.ColWidth(4) = 0
'                MS.ColWidth(0) = 2500
'                MS.ColWidth(1) = 1100
'                MS.ColWidth(2) = 1800
'                MS.TextMatrix(0, 0) = "Product Name"
'                MS.TextMatrix(0, 1) = "Quantity"
'                MS.ColAlignment(1) = 7
'                SCl MS, 1
'                MS.TextMatrix(0, 2) = "Discount Amt Per Case"
'                SCl MS, 2
'                MS.ColAlignment(2) = 7
'                If RSINV.State = 1 Then RSINV.Close
'                RSINV.Open "select sum(tot_pcs) as tot_pcs,sub_unt,prd_idy from inv where inv_idy in(select inv_idy from inl where inv_dat>=ctod('" & fdate.Value & "') and inv_dat<=ctod('" & tdate.Value & "') and cst_idy='" & Trim(retid) & "') AND stk_typ in('Salable','Free') group by prd_idy", CON, adOpenKeyset, adLockOptimistic
'                While Not RS.EOF
'                    If Not RSINV.EOF Then RSINV.MoveFirst
'                    MS.Rows = MS.Rows + 1
'                    MST.Rows = MST.Rows + 1
'                    If RSPRD.State = 1 Then RSPRD.Close
'                    RSPRD.Open "select prd_nme,SUB_UNT from prd where prd_idy='" & RS("prd_idy") & "'", CON, adOpenKeyset, adLockOptimistic
'                    MS.TextMatrix(R, 0) = RSPRD(0)
'                    MST.TextMatrix(R, 0) = RSPRD(0)
'                    MS.TextMatrix(R, 4) = RSPRD(1)
'                    MST.TextMatrix(R, 6) = RSPRD(1)
'                    MS.TextMatrix(R, 3) = RS("prd_idy")
'                    MST.TextMatrix(R, 7) = RS("prd_idy")
'                    MS.TextMatrix(R, 1) = QTY(RS("prd_qty"))
'                    MST.TextMatrix(R, 1) = QTY(RS("prd_qty"))
'                    MS.TextMatrix(R, 2) = convert(RS("dis_amt"))
'                    MST.TextMatrix(R, 2) = convert(RS("DIS_AMT"))
'                    Do While Not RSINV.EOF
'                        If RS("prd_idy") = RSINV("prd_idy") Then
'                            MST.TextMatrix(R, 3) = STM(Val(RSINV(0)), Val(RSINV(1)))
'                            'MST.TextMatrix(r, 4) = convert(CASEAMT(Trim(MS.TextMatrix(r, 3)), Trim(MS.TextMatrix(r, 4)), Trim(MS.TextMatrix(r, 2))))
'                            Exit Do
'                        End If
'                        RSINV.MoveNext
'                    Loop
'                    R = R + 1
'                    RS.MoveNext
'                Wend
'                If RS.RecordCount > 0 Then
'                RS.MoveFirst
'                For d = 0 To 11
'                    V = d
'                    If Val(RS(d + 20)) > 0 Then
'                        If d = 0 Then
'                            pdate(d).Enabled = True
'                            pdate(d).Value = RS(d + 8)
'                            Text2(d) = convert(RS(d + 20))
'                            'Text3(d) = Val(Qpsamt.Text) - Val(Text2(d))
'                        Else
'                            pdate(d).Enabled = True
'                            pdate(d).Value = RS(d + 8)
'                            Text2(d) = convert(RS(d + 20))
'                            'Text3(d) = Val(Text3(d - 1)) - Val(Text2(d))
'                        End If
'                    Else
'                        Exit For
'                    End If
'                Next
'                End If
'                For R = 1 To MST.Rows - 1
'                    If Val(MST.TextMatrix(R, 1)) >= Val(MST.TextMatrix(R, 3)) Then
'                        Exit For
'                    End If
'                Next
'                If R = MST.Rows Then
'                    For R = 1 To MST.Rows - 1
'                        MST.TextMatrix(R, 4) = convert(CASEAMT(Trim(MST.TextMatrix(R, 3)), Trim(MS.TextMatrix(R, 6)), Trim(MS.TextMatrix(R, 2))))
'                    Next
'                End If
'                    'smnid.Enabled = False
'                MS.SetFocus
'                MS.Col = 2
'            End If
'            retid.SetFocus
'            SendKeys "{HOME}+{END}"
'            MD = True
'        End If
'    Else
'        qpslist
'    End If
'End If
'End Sub

'Private Sub qpslist()
'If CHECKTB("qps.dbf") = 1 Then
'    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'    Exit Sub
'End If
'Dim ls As ListItem
'Dim RS As New ADODB.Recordset
'If RS.State = 1 Then RS.Close
'RS.Open "select * from qps where txn_idy like '" & Trim(txnid) & "%' group by txn_idy", CON, adOpenStatic
'Label20.Caption = "Select Transaction"
'lv.ColumnHeaders.Clear
'lv.ColumnHeaders.Add , , "Transaction ID", 2000
'lv.ColumnHeaders.Add , , "Retailer", 2500
''lv.ColumnHeaders.Add , , "Commission Amt", 1500, 2
'lv.ListItems.Clear
'If RS.RecordCount > 0 Then
'    For A = 1 To RS.RecordCount
'        Set ls = lv.ListItems.Add(, , Trim(RS("txn_idy")))
'        If rsact.State = 1 Then rsact.Close
'        rsact.Open "select act_nme from act where act_idy='" & RS("act_idy") & "'", CON
'        If Not rsact.EOF Then
'            ls.ListSubItems.Add , , Trim(rsact("act_nme"))
'        End If
' '       ls.ListSubItems.Add , , Trim(rs("com_amt") & "")
'        RS.MoveNext
'    Next
'    SSTab1.TabEnabled(3) = True
'    SSTab1.Tab = 3
'    lv.Visible = True
'    lv.SetFocus
'Else
'    If SSTab1.Tab = 0 Then
'        MsgBox "Transaction ID Not Found, Define Transaction !", vbExclamation, "Effmcg"
'        SSTab1.Tab = 0
'        txnid.Enabled = True
'        txnid.SetFocus
'        MD = False
'    End If
'End If
'End Sub

Private Sub tog(b As Boolean)
txtDisamt.Enabled = b
txtQty.Enabled = b
End Sub

Private Sub txtDisamt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Val(txtDisamt) > 0 Then
        txtDisamt = convert(txtDisamt)
        saverecord
    Else
        MsgBox "Enter Product Discount Per Case", vbExclamation
        txtDisamt.SetFocus
        SendKeys "{home}+{end}"
    End If
End If
End Sub
Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Val(txtQty) > 0 Then
        txtDisamt.Enabled = True
        txtDisamt.SetFocus
        SendKeys "{home}+{end}"
    Else
        MsgBox "Enter All Products Quantity", vbExclamation
        txtQty.SetFocus
        SendKeys "{home}+{end}"
    End If
End If
End Sub
Private Sub NewRecord()
MD = False
cancel
DISPPH
retid.SetFocus
End Sub

Private Sub cancel()
clrctr Me
'txnid = ""
retid.Enabled = True
retid = ""
retname.Caption = ""
fdate.Value = Date
tdate.Value = Date
All.Enabled = True
All.Value = 0
txtQty = "0.000"
txtDisamt = "0.00"
MS.Clear
txtvar = ""
txtvar.Visible = False
MS.Rows = 1
If SSTAB1.Tab = 3 Then
    lv.Visible = False
    SSTAB1.Tab = 0
End If
'txnid.Enabled = True
retid.SetFocus
DISPPH
End Sub

Private Sub saverecord()
If MsgBox("Save Record ?", vbYesNo + vbExclamation) = vbNo Then
    cancel
    Exit Sub
End If
If Trim(retid) = "" Then
    MsgBox "Select a Retailer ID", vbExclamation
    retid.SetFocus
    Exit Sub
End If
If MD = False Then
    If rsqps.State = 1 Then rsqps.Close
    rsqps.Open "select * from qps where txn_idy=''", CON, adOpenKeyset, adLockOptimistic
    R = 1
    If All.Value = 1 Then
'        txnid = increment("idy_qps")
        rsqps.AddNew
        For K = 0 To rsqps.Fields.Count - 1
            If rsqps(K).Type = adBigInt Or rsqps(K).Type = adNumeric Or rsqps(K).Type = adInteger Or rsqps(K).Type = adSingle Or rsqps(K).Type = adDouble Then
                rsqps(K) = 0
            ElseIf rsqps(K).Type = adDate Or rsqps(K).Type = adDBDate Or rsqps(K).Type = adDBTimeStamp Then
                rsqps(K) = Date
            End If
        Next
        'rsqps("txn_idy") = txnid
        rsqps("act_idy") = Trim(retid)
        rsqps("frm_dat") = fdate.Value
        rsqps("to_dat") = tdate.Value
        rsqps("chk_flg") = "A"
        rsqps("dis_amt") = Val(txtDisamt)
        rsqps("prd_qty") = Val(txtQty)
        rsqps.Update
        'MsgBox "New Transaction ID is " & txnid, vbExclamation
        cancel
    Else
        S = 0
        For R = 1 To MS.Rows - 1
            If Val(MS.TextMatrix(R, 1)) > 0 And Val(MS.TextMatrix(R, 2)) > 0 Then
                S = 1
                Exit For
            End If
        Next
        If S = 1 Then
'            txnid = increment("idy_qps")
            For R = 1 To MS.Rows - 1
                If Val(MS.TextMatrix(R, 1)) > 0 And Val(MS.TextMatrix(R, 2)) > 0 Then
                    S = 1
                    rsqps.AddNew
                    For K = 0 To rsqps.Fields.Count - 1
                        If rsqps(K).Type = adBigInt Or rsqps(K).Type = adNumeric Or rsqps(K).Type = adInteger Or rsqps(K).Type = adSingle Or rsqps(K).Type = adDouble Then
                            rsqps(K) = 0
                        ElseIf rsqps(K).Type = adDate Or rsqps(K).Type = adDBDate Or rsqps(K).Type = adDBTimeStamp Then
                            rsqps(K) = Date
                        End If
                    Next
                    'rsqps("txn_idy") = txnid
                    rsqps("act_idy") = Trim(retid)
                    rsqps("frm_dat") = fdate.Value
                    rsqps("to_dat") = tdate.Value
                    rsqps("chk_flg") = "S"
                    rsqps("dis_amt") = Val(MS.TextMatrix(R, 2))
                    rsqps("prd_idy") = Trim(MS.TextMatrix(R, 3))
                    rsqps("prd_qty") = Val(MS.TextMatrix(R, 1))
                    rsqps.Update
                End If
            Next
            'MsgBox "New Transaction ID is " & txnid, vbExclamation
            cancel
        Else
            MsgBox "Enter Discount amount per case", vbExclamation
            MS.SetFocus
            MS.Col = 2
            MS.Row = 1
        End If
    End If
Else
    If rsqps.State = 1 Then rsqps.Close
    'rsqps.Open "SELECT * FROM qps WHERE TXN_IDY='" & Trim(txnid) & "'", CON, adOpenKeyset, adLockOptimistic
    If All.Value = 1 Then
        rsqps("act_idy") = Trim(retid)
        rsqps("frm_dat") = fdate.Value
        rsqps("to_dat") = tdate.Value
        rsqps("chk_flg") = "A"
        rsqps("prd_qty") = Val(txtQty)
        rsqps("dis_amt") = Val(txtDisamt)
        For d = 0 To 11
            If pdate(d).Enabled = True Then 'And Val(Text2(d)) > 0 Then
                rsqps(d + 8) = pdate(d).Value
                rsqps(d + 20) = convert(Val(Text2(d).Text))
            Else
                Exit For
            End If
        Next
        rsqps.Update
        cancel
    Else
        R = 1
        While Not rsqps.EOF
            rsqps("act_idy") = Trim(retid)
            rsqps("frm_dat") = fdate.Value
            rsqps("to_dat") = tdate.Value
            rsqps("chk_flg") = "S"
            rsqps("dis_amt") = Val(MS.TextMatrix(R, 2))
            rsqps("prd_idy") = MS.TextMatrix(R, 3)
            rsqps("prd_qty") = Val(MS.TextMatrix(R, 1))
            For d = 0 To 11
                If pdate(d).Enabled = True Then 'And amt1(d).Enabled = True And bal1(d).Enabled = True Then
                    rsqps(d + 8) = pdate(d).Value
                    rsqps(d + 20) = convert(Val(Text2(d).Text))
                Else
                    Exit For
                End If
            Next
            rsqps.Update
            rsqps.MoveNext
            R = R + 1
        Wend
        cancel
    End If
End If
End Sub

Private Sub dispprd()
If RSPRD.State = 1 Then RSPRD.Close
RSPRD.Open "select prd_idy,prd_nme,sub_unt from prd", CON, adOpenKeyset, adLockOptimistic
R = 1
DISPPH
While Not RSPRD.EOF
    MS.Rows = MS.Rows + 1
    MS.TextMatrix(R, 0) = RSPRD(1)
    MS.TextMatrix(R, 3) = RSPRD(0)
    MS.TextMatrix(R, 4) = RSPRD(2)
    MS.TextMatrix(R, 1) = "0.000"
    MS.TextMatrix(R, 2) = "0.00"
    R = R + 1
    RSPRD.MoveNext
Wend
End Sub

Private Sub TXTSAL_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    If MST.Col = 3 Then
'        If Val(CMP1()) > Val(MST.TextMatrix(MST.Row, 6)) Then
'            MsgBox "Sub Units are " & MST.TextMatrix(MS.Row, 6)
'            TXTSAL.SetFocus
'            SendKeys "{home}+{end}"
'        ElseIf Val(TXTSAL) > Val(MST.TextMatrix(MST.Row, 1)) Then
'            MsgBox "Invalid Stock", vbExclamation
'            TXTSAL.SetFocus
'            SendKeys "{home}+{end}"
'        Else
'            MST.Text = QTY(Val(Trim(TXTSAL)))
'            MST.TextMatrix(MST.Row, 4) = convert(CASEAMT(Trim(MST.Text), Trim(MST.TextMatrix(MST.Row, 6)), Trim(MS.TextMatrix(MST.Row, 2))))
'            TXTSAL.Visible = False
'            If MST.Row <> MST.Rows - 1 Then
'                MST.Row = MST.Row + 1
'                MST.Col = 3
'            End If
'            MST.SetFocus
'        End If
'    End If
'End If
End Sub

Private Sub txtvar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MS.Col = 1 Then
        RT = Val(CMP())
        If RT >= Val(MS.TextMatrix(MS.Row, 4)) Then
            MsgBox "Sub Units are " & Val(MS.TextMatrix(MS.Row, 4)), vbExclamation
            txtvar.Visible = True
            txtvar.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        ElseIf RT = 0 Then
            MS.Text = Trim(Val(txtvar)) & ".000"
            txtvar = ""
            txtvar.Visible = False
            MS.Col = MS.Col + 1
        Else
            MS.Text = QTY(Trim(Val(txtvar)))
            txtvar = ""
            txtvar.Visible = False
            MS.Col = MS.Col + 1
        End If
        MS.SetFocus
    ElseIf MS.Col = 2 Then
        MS.Text = convert(txtvar)
        txtvar = ""
        txtvar.Visible = False
        If MS.Row <> MS.Rows - 1 Then
            MS.Row = MS.Row + 1
            MS.Col = 1
        End If
        MS.SetFocus
    End If
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
Else
    CMP = 0
End If
End Function

Private Sub DISPT()
With MST
    .Rows = 1
    .Cols = 9
    .TextMatrix(0, 0) = "Product Name"
    .ColWidth(0) = 2500
    .TextMatrix(0, 1) = "Target Qty"
    .ColWidth(1) = 1000
    .TextMatrix(0, 2) = "Target Amt."
    .ColWidth(2) = 1000
    .TextMatrix(0, 3) = "Sale Qty"
    .ColWidth(3) = 1000
    .TextMatrix(0, 4) = "QPS/UPS Amt."
    .ColWidth(4) = 1250
    .TextMatrix(0, 5) = "Paid Amt."
    .ColWidth(5) = 0
    .TextMatrix(0, 6) = "Sb Units"
    .ColWidth(6) = 0
    .TextMatrix(0, 7) = "Product Id"
    .ColWidth(7) = 0
    .TextMatrix(0, 8) = "Balance"
    .ColWidth(8) = 0
End With
End Sub
'Public Function CMP1()
'Dim chk As Integer, CHK1 As String
'chk = InStr(Trim(STR(Val(TXTSAL))), ".")
'If chk <> 0 Then
'    CHK1 = Mid(Trim(STR(Val(TXTSAL))), chk + 1, Len(Trim(Val(TXTSAL))))
'    If Len(Trim(CHK1)) = 1 Then
'        CMP1 = Val(CHK1) * 100
'    ElseIf Len(Trim(CHK1)) = 2 Then
'        CMP1 = Val(CHK1) * 10
'    ElseIf Len(Trim(CHK1)) = 3 Then
'        CMP1 = Val(CHK1) * 1
'    Else
'        CMP1 = Val(Mid(Trim(CHK1), 1, 3)) * 1
'    End If
'Else
'    CMP1 = 0
'End If
'End Function
Private Sub DISPPH()
MS.Rows = 1
MS.Cols = 5
MS.ColWidth(3) = 0
MS.ColWidth(4) = 0
MS.ColWidth(0) = 2500
MS.ColWidth(1) = 1100
MS.ColWidth(2) = 2000
MS.TextMatrix(0, 0) = "Product Name"
MS.TextMatrix(0, 1) = "Quantity"
MS.ColAlignment(1) = 7
SCl MS, 1
MS.TextMatrix(0, 2) = "Discount Amt Per Case"
SCl MS, 2
MS.ColAlignment(2) = 7
End Sub
