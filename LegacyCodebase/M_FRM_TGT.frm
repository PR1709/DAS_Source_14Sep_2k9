VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_TGT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Targets - Salesmen"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "M_FRM_TGT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4933.914
   ScaleMode       =   0  'User
   ScaleWidth      =   8970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3720
      TabIndex        =   44
      Top             =   840
      Width           =   3495
      Begin VB.Label M_SMN_NME 
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
         Height          =   340
         Left            =   240
         TabIndex        =   45
         Top             =   120
         Width           =   3165
      End
   End
   Begin VB.TextBox smnid 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   337
      Left            =   2564
      TabIndex        =   43
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   41
      Top             =   960
      Width           =   1095
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Salesman ID"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   60
         Width           =   900
      End
   End
   Begin TabDlg.SSTab SSTAB1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Product (SKU) Level"
      TabPicture(0)   =   "M_FRM_TGT.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MS"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TGTQTY"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Brand Level"
      TabPicture(1)   =   "M_FRM_TGT.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "msb"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(4)=   "B_TGTQTY"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "&Segment Level"
      TabPicture(2)   =   "M_FRM_TGT.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(1)=   "mss"
      Tab(2).Control(2)=   "Frame5"
      Tab(2).Control(3)=   "Frame6"
      Tab(2).Control(4)=   "S_TGTQTY"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "&Category Level"
      TabPicture(3)   =   "M_FRM_TGT.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label9"
      Tab(3).Control(1)=   "msc"
      Tab(3).Control(2)=   "Frame7"
      Tab(3).Control(3)=   "Frame8"
      Tab(3).Control(4)=   "C_TGTQTY"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "View"
      TabPicture(4)   =   "M_FRM_TGT.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label10"
      Tab(4).Control(1)=   "list1"
      Tab(4).ControlCount=   2
      Begin VB.TextBox B_TGTQTY 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   -73560
         MaxLength       =   11
         TabIndex        =   36
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox C_TGTQTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   -73320
         MaxLength       =   11
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox S_TGTQTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   -73320
         MaxLength       =   11
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame8 
         Caption         =   "FMCG Calender Dates"
         Height          =   735
         Left            =   -72440
         TabIndex        =   27
         Top             =   885
         Width           =   4575
         Begin VB.Label C_START 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   600
            TabIndex        =   31
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label C_END 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2760
            TabIndex        =   30
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "End"
            Height          =   195
            Left            =   2400
            TabIndex        =   29
            Top             =   360
            Width           =   285
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Start"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   330
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   -67800
         TabIndex        =   25
         Top             =   900
         Width           =   1575
         Begin VB.CommandButton C_CMD 
            Caption         =   "G&et Category"
            Enabled         =   0   'False
            Height          =   435
            Left            =   120
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   200
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "FMCG Calender Dates"
         Height          =   735
         Left            =   -72440
         TabIndex        =   19
         Top             =   885
         Width           =   4575
         Begin VB.Label S_START 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   600
            TabIndex        =   23
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label S_END 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2760
            TabIndex        =   22
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "End"
            Height          =   195
            Left            =   2400
            TabIndex        =   21
            Top             =   360
            Width           =   285
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Start"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   330
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1320
         Left            =   -67800
         TabIndex        =   17
         Top             =   320
         Width           =   1575
         Begin VB.CommandButton M_CAT 
            Caption         =   "Cate&gory ++"
            Enabled         =   0   'False
            Height          =   435
            Left            =   120
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton S_CMD 
            Caption         =   "G&et Segment"
            Enabled         =   0   'False
            Height          =   435
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   800
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "FMCG Calender Dates"
         Height          =   735
         Left            =   -72440
         TabIndex        =   12
         Top             =   885
         Width           =   4575
         Begin VB.Label B_START 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   600
            TabIndex        =   16
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label B_END 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2760
            TabIndex        =   15
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "End"
            Height          =   195
            Left            =   2400
            TabIndex        =   14
            Top             =   360
            Width           =   285
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Start"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   330
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -67800
         TabIndex        =   10
         Top             =   915
         Width           =   1575
         Begin VB.CommandButton B_CMD 
            Caption         =   "G&et Brands"
            Enabled         =   0   'False
            Height          =   435
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   195
            Width           =   1335
         End
      End
      Begin VB.TextBox TGTQTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   5400
         MaxLength       =   11
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "FMCG Calender Dates"
         Height          =   735
         Left            =   2560
         TabIndex        =   3
         Top             =   885
         Width           =   4575
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Start"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "End"
            Height          =   195
            Left            =   2400
            TabIndex        =   6
            Top             =   360
            Width           =   285
         End
         Begin VB.Label ENDDATE 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2760
            TabIndex        =   5
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label Start 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   600
            TabIndex        =   4
            Top             =   285
            Width           =   1620
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1320
         Left            =   7200
         TabIndex        =   1
         Top             =   320
         Width           =   1575
         Begin VB.CommandButton M_SEG 
            Caption         =   "S&egment ++"
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   170
            Width           =   1335
         End
         Begin VB.CommandButton M_BRD 
            Caption         =   "B&rand ++"
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   542
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Get SKUs"
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   915
            Width           =   1335
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   3090
         Left            =   60
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1680
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   5450
         _Version        =   393216
         Rows            =   3
         Cols            =   106
         FixedRows       =   2
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   12582912
         FocusRect       =   2
         GridLines       =   2
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid mss 
         Height          =   3090
         Left            =   -74940
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5450
         _Version        =   393216
         Rows            =   3
         Cols            =   105
         FixedRows       =   2
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   12582912
         FocusRect       =   2
         GridLines       =   2
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid msc 
         Height          =   3090
         Left            =   -74940
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5450
         _Version        =   393216
         Rows            =   3
         Cols            =   105
         FixedRows       =   2
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   12582912
         FocusRect       =   2
         GridLines       =   2
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid msb 
         Height          =   3090
         Left            =   -74940
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5450
         _Version        =   393216
         Rows            =   3
         Cols            =   105
         FixedRows       =   2
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   12582912
         FocusRect       =   2
         GridLines       =   2
         AllowUserResizing=   1
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   14208
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
            Text            =   "Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Categories Please Wait....."
         Height          =   195
         Left            =   -74880
         TabIndex        =   51
         Top             =   4800
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Segment's Please Wait....."
         Height          =   195
         Left            =   -74880
         TabIndex        =   50
         Top             =   4800
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Brand's Please Wait....."
         Height          =   195
         Left            =   -74880
         TabIndex        =   49
         Top             =   4800
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading SKU's Please Wait....."
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   4800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label10 
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
         TabIndex        =   47
         Top             =   600
         Width           =   8055
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
            Picture         =   "M_FRM_TGT.frx":04CE
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":0DDA
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":112E
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":1766
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":1ABA
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":1F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":220E
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":261A
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":2726
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":2A7A
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":2E8E
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TGT.frx":33DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
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
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SR"
                  Text            =   "Target Analysis Report"
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
            Object.Visible         =   0   'False
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
Attribute VB_Name = "M_FRM_TGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WEK As New ADODB.Recordset
Dim Cls1 As New Class1, A As Integer

Private Sub m_cat_Click()
Cls1.MRG_CAT
End Sub

Private Sub B_CMD_Click()
Label7.Visible = True
Cls1.TGT_LOD_BRD
B_CMD.Enabled = False
msb.SetFocus
msb.Col = 1
msb.Row = 2
Label7.Visible = False
End Sub

Private Sub B_ROU_LST_Click()
If CHECKTB("mar.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
    Dim rs As New ADODB.Recordset
    If B_ROU_LST = "All" Then
        rs.Open "select * from mar", CON, adOpenKeyset
    Else
        rs.Open "select * from mar where rou_idy like '" & Trim(B_ROU_LST) & "'", CON, adOpenKeyset
    End If
    LIST2.CLEAR
    For A = 1 To rs.RecordCount
        LIST2.AddItem Trim(rs!mar_nme)
        rs.MoveNext
    Next
End Sub

Private Sub B_ROU_LST_KeyPress(K As Integer)
If K = 13 Then
    B_ROU_LST.Enabled = False
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    B_CMD.Enabled = True
    B_CMD.SetFocus
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub B_TGTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    msb.Text = Int(Val(B_TGTQTY))
    If msb.Col <> 104 Then msb.Col = msb.Col + 1
    B_TGTQTY.Visible = False
    msb.SetFocus
End If
End Sub

Private Sub B_TGTQTY_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub C_CMD_Click()
Label9.Visible = True
Cls1.TGT_LOD_CAT
C_CMD.Enabled = False
MSC.SetFocus
MSC.Col = 1
MSC.Row = 2
Label9.Visible = False
End Sub

Private Sub C_TGTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSC.Text = Int(Val(C_TGTQTY))
    If MSC.Col <> 104 Then MSC.Col = MSC.Col + 1
    C_TGTQTY.Visible = False
    MSC.SetFocus
End If
End Sub

Private Sub C_TGTQTY_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_BRD_Click()
Cls1.MRG_BRD
End Sub

Private Sub M_SEG_Click()
Cls1.MRG_SEG
End Sub

Private Sub MS_Scroll()
If TGTQTY.Visible = True Then
    TGTQTY = ""
    TGTQTY.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub msb_Scroll()
If B_TGTQTY.Visible = True Then
    B_TGTQTY = ""
    B_TGTQTY.Visible = False
    msb.SetFocus
End If
End Sub

Private Sub msc_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
Exit Sub
End If
If K <> 27 Then
    If MSS.Col <> 0 Then
        C_TGTQTY.Visible = True
        C_TGTQTY.Top = MSC.Top + MSC.CellTop
        C_TGTQTY.Left = MSC.CellLeft + MSC.Left
        C_TGTQTY.Width = MSC.CellWidth
        C_TGTQTY.Height = MSC.CellHeight
        C_TGTQTY.Visible = True
        C_TGTQTY.SetFocus
        If K <> 13 Then C_TGTQTY = Chr(K)
        C_TGTQTY.SelStart = Len(C_TGTQTY)
    End If
End If
End Sub

Private Sub msc_Scroll()
If C_TGTQTY.Visible = True Then
    C_TGTQTY = ""
    C_TGTQTY.Visible = False
    MSC.SetFocus
End If
End Sub

Private Sub MSS_Scroll()
If S_TGTQTY.Visible = True Then
    S_TGTQTY = ""
    S_TGTQTY.Visible = False
    MSS.SetFocus
End If
End Sub

Private Sub S_CMD_Click()
Label8.Visible = True
Cls1.TGT_LOD_SEG
S_CMD.Enabled = False
MSS.SetFocus
MSS.Col = 1
MSS.Row = 2
Label8.Visible = False
End Sub

Private Sub S_TGTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSS.Text = Int(Val(S_TGTQTY))
    If MSS.Col <> 104 Then MSS.Col = MSS.Col + 1
    S_TGTQTY.Visible = False
    MSS.SetFocus
End If
End Sub

Private Sub S_TGTQTY_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub C_ROU_LST_Click()
If CHECKTB("mar.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If
    Dim rs As New ADODB.Recordset
    If C_ROU_LST = "All" Then
        rs.Open "select * from mar", CON, adOpenKeyset
    Else
        rs.Open "select * from mar where rou_idy like '" & Trim(C_ROU_LST) & "'", CON, adOpenKeyset
    End If
    list4.CLEAR
    For A = 1 To rs.RecordCount
        list4.AddItem Trim(rs!mar_nme)
        rs.MoveNext
    Next
End Sub

Private Sub C_ROU_LST_KeyPress(K As Integer)
If K = 13 Then
    C_ROU_LST.Enabled = False
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(1) = False
    C_CMD.Enabled = True
    C_CMD.SetFocus
End If
End Sub

Private Sub Command3_Click()
Command3.Enabled = False
Label2.Visible = True
Label2.Refresh
Cls1.TGT_LOAD_SKU
MS.SetFocus
MS.Col = 1
Label2.Visible = False
End Sub

Private Sub Form_Activate()
'If SSTab1.TabEnabled(0) = True Then
'    SSTab1.Tab = 0
'    If ROU_LST.Enabled = True Then ROU_LST.SetFocus
'ElseIf SSTab1.TabEnabled(1) = True Then
'    SSTab1.Tab = 1
'    If B_ROU_LST.Enabled = True Then B_ROU_LST.SetFocus
'ElseIf SSTab1.TabEnabled(2) = True Then
'    SSTab1.Tab = 2
'    If S_ROU_LST.Enabled = True Then S_ROU_LST.SetFocus
'ElseIf SSTab1.TabEnabled(3) = True Then
'    SSTab1.Tab = 3
'    If C_ROU_LST.Enabled = True Then C_ROU_LST.SetFocus
'End If
SSTab1.Tab = 0
If smnid.Enabled = True Then smnid.SetFocus
MS.TextMatrix(0, 0) = "Product (SKU) Name"
msb.TextMatrix(0, 0) = "Brand"
MSS.TextMatrix(0, 0) = "Product Segment"
MSC.TextMatrix(0, 0) = "Product Category"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 80 And Shift = 2 Then PRINTRECORD
    If KeyCode = 83 And Shift = 2 Then Cls1.TGT_SAVE
    If KeyCode = 88 And Shift = 2 Then Unload Me
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 69 And Shift = 2 Then smnlist
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If smnid.Enabled = True Then
        Unload Me
    Else
        Cls1.TGT_SAVE
    End If
End If
End Sub

Private Sub Form_Load()
If CHECKTB("wek.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
MS.MergeCells = flexMergeRestrictAll
MS.MergeRow(0) = True
MS.ColWidth(0) = 2500
MS.ColWidth(105) = 0
msb.MergeCells = flexMergeRestrictAll
msb.MergeRow(0) = True
msb.ColWidth(0) = 2500
MSC.MergeCells = flexMergeRestrictAll
MSC.MergeRow(0) = True
MSC.ColWidth(0) = 2500
MSS.MergeCells = flexMergeRestrictAll
MSS.MergeRow(0) = True
MSS.ColWidth(0) = 2500

b = 1
For A = 1 To 104 Step 2
    MS.TextMatrix(0, A) = "W" & b
    MS.Row = 0
    MS.Col = A
    MS.CellAlignment = 4
    MS.TextMatrix(0, A + 1) = "W" & b
    MS.Row = 0
    MS.Col = A + 1
    MS.CellAlignment = 4
    b = b + 1
Next

For A = 1 To 104
    If A Mod 2 = 0 Then
        MS.TextMatrix(1, A) = "Qty"
        MS.ColAlignment(A) = 7
        SCl MS, A, 1
    Else
        MS.TextMatrix(1, A) = "Vol."
        MS.ColAlignment(A) = 7
        SCl MS, A, 1
    End If
    MS.ColWidth(A) = 600
Next

b = 1
For A = 1 To 104 Step 2
    msb.TextMatrix(0, A) = "W" & b
    msb.Row = 0
    msb.Col = A
    msb.CellAlignment = 4
    msb.TextMatrix(0, A + 1) = "W" & b
    msb.Row = 0
    msb.Col = A + 1
    msb.CellAlignment = 4
    b = b + 1
Next

For A = 1 To 104
    If A Mod 2 = 0 Then
        msb.TextMatrix(1, A) = "Qty"
        msb.ColAlignment(A) = 7
        SCl msb, A, 1
    Else
        msb.TextMatrix(1, A) = "Vol."
        msb.ColAlignment(A) = 7
        SCl msb, A, 1
    End If
    msb.ColWidth(A) = 600
Next

b = 1
For A = 1 To 104 Step 2
    MSS.TextMatrix(0, A) = "W" & b
    MSS.Row = 0
    MSS.Col = A
    MSS.CellAlignment = 4
    MSS.TextMatrix(0, A + 1) = "W" & b
    MSS.Row = 0
    MSS.Col = A + 1
    MSS.CellAlignment = 4
    b = b + 1
Next

For A = 1 To 104
    If A Mod 2 = 0 Then
        MSS.TextMatrix(1, A) = "Qty"
        MSS.ColAlignment(A) = 7
        SCl MSS, A, 1
    Else
        MSS.TextMatrix(1, A) = "Vol."
        MSS.ColAlignment(A) = 7
        SCl MSS, A, 1
    End If
    MSS.ColWidth(A) = 600
Next

b = 1
For A = 1 To 104 Step 2
    MSC.TextMatrix(0, A) = "W" & b
    MSC.Row = 0
    MSC.Col = A
    MSC.CellAlignment = 4
    MSC.TextMatrix(0, A + 1) = "W" & b
    MSC.Row = 0
    MSC.Col = A + 1
    MSC.CellAlignment = 4
    b = b + 1
Next

For A = 1 To 104
    If A Mod 2 = 0 Then
        MSC.TextMatrix(1, A) = "Qty"
        MSC.ColAlignment(A) = 7
        SCl MSC, A, 1
    Else
        MSC.TextMatrix(1, A) = "Vol."
        MSC.ColAlignment(A) = 7
        SCl MSC, A, 1
    End If
    MSC.ColWidth(A) = 600
Next
MS.ColAlignment(0) = 1
msb.ColAlignment(0) = 1
MSS.ColAlignment(0) = 1
MSC.ColAlignment(0) = 1
SSTab1.TabEnabled(4) = False
If WEK.State = 1 Then WEK.Close
WEK.Open "SELECT * FROM WEK", CON, adOpenKeyset
End Sub

Private Sub MS_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
Exit Sub
End If
If K <> 27 Then
    If MS.Col <> 0 And MS.Col <> 53 Then
        TGTQTY.Visible = True
        TGTQTY.Top = MS.Top + MS.CellTop
        TGTQTY.Left = MS.CellLeft + MS.Left
        TGTQTY.Width = MS.CellWidth
        TGTQTY.Height = MS.CellHeight
        TGTQTY.Visible = True
        TGTQTY.SetFocus
        If K <> 13 Then TGTQTY = Chr(K)
        TGTQTY.SelStart = Len(TGTQTY)
    End If
End If
End Sub

Private Sub MS_LeaveCell()
TGTQTY.Visible = False
End Sub

Private Sub MS_RowColChange()
If MS.Rows > 2 Then
If WEK.State = 1 Then
If WEK.RecordCount > 0 And Trim(MS.TextMatrix(2, 0)) <> "" Then
    WEK.Filter = "WEK_IDY=" & Val(Mid(Trim(MS.TextMatrix(0, MS.Col)), 2, 2))
    If WEK.RecordCount > 0 Then
    Start = Format(WEK!FRM_DAT, "DD/MMM/YYYY")
    ENDDATE = Format(WEK!TO_DAT, "DD/MMM/YYYY")
    End If
End If
End If
End If
End Sub

Private Sub msb_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
Exit Sub
End If
If K <> 27 Then
    If msb.Col <> 0 Then
        B_TGTQTY.Visible = True
        B_TGTQTY.Top = msb.Top + msb.CellTop
        B_TGTQTY.Left = msb.CellLeft + msb.Left
        B_TGTQTY.Width = msb.CellWidth
        B_TGTQTY.Height = msb.CellHeight
        B_TGTQTY.Visible = True
        B_TGTQTY.SetFocus
        If K <> 13 Then B_TGTQTY = Chr(K)
        B_TGTQTY.SelStart = Len(B_TGTQTY)
    End If
End If
End Sub

Private Sub msb_RowColChange()
If msb.Rows > 2 Then
    If WEK.State = 1 Then
        If WEK.RecordCount > 0 And Trim(msb.TextMatrix(2, 0)) <> "" Then
            WEK.Filter = "WEK_IDY=" & Val(Mid(Trim(msb.TextMatrix(0, msb.Col)), 2, 2))
            If WEK.RecordCount > 0 Then
            B_START = Format(WEK!FRM_DAT, "DD/MMM/YYYY")
            B_END = Format(WEK!TO_DAT, "DD/MMM/YYYY")
            End If
        End If
    End If
End If
End Sub
Private Sub msC_RowColChange()
If MSC.Rows > 2 Then
    If WEK.State = 1 Then
        If WEK.RecordCount > 0 And Trim(MSC.TextMatrix(2, 0)) <> "" Then
            WEK.Filter = "WEK_IDY=" & Val(Mid(Trim(MSC.TextMatrix(0, MSC.Col)), 2, 2))
            If WEK.RecordCount > 0 Then
            C_START = Format(WEK!FRM_DAT, "DD/MMM/YYYY")
            C_END = Format(WEK!TO_DAT, "DD/MMM/YYYY")
            End If
        End If
    End If
End If
End Sub

Private Sub MSS_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
Exit Sub
End If
If K <> 27 Then
    If MSS.Col <> 0 Then
        S_TGTQTY.Visible = True
        S_TGTQTY.Top = MSS.Top + MSS.CellTop
        S_TGTQTY.Left = MSS.CellLeft + MSS.Left
        S_TGTQTY.Width = MSS.CellWidth
        S_TGTQTY.Height = MSS.CellHeight
        S_TGTQTY.Visible = True
        S_TGTQTY.SetFocus
        If K <> 13 Then S_TGTQTY = Chr(K)
        S_TGTQTY.SelStart = Len(S_TGTQTY)
    End If
End If
End Sub

Private Sub msS_RowColChange()
If MSS.Rows > 2 Then
    If WEK.State = 1 Then
        If WEK.RecordCount > 0 And Trim(MSS.TextMatrix(2, 0)) <> "" Then
            WEK.Filter = "WEK_IDY=" & Val(Mid(Trim(MSS.TextMatrix(0, MSS.Col)), 2, 2))
            If WEK.RecordCount > 0 Then
            S_START = Format(WEK!FRM_DAT, "DD/MMM/YYYY")
            S_END = Format(WEK!TO_DAT, "DD/MMM/YYYY")
            End If
        End If
    End If
End If
End Sub
'Private Sub ROU_LST_Click()
'If CHECKTB("mar.dbf") = 1 Then
'        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'        cancelrecord
'        Exit Sub
'End If
'
'    Dim rs As New ADODB.Recordset
'    If ROU_LST = "All" Then
'        rs.Open "select * from mar", CON, adOpenKeyset
'    Else
'        rs.Open "select * from mar where rou_idy like '" & Trim(ROU_LST) & "'", CON, adOpenKeyset
'    End If
'    list1.Clear
'    For A = 1 To rs.RecordCount
'        list1.AddItem Trim(rs!mar_nme)
'        rs.MoveNext
'    Next
'End Sub

'Private Sub ROU_LST_kEYPRESS(K As Integer)
'If K = 13 Then
'    ROU_LST.Enabled = False
'    sstab1.TabEnabled(1) = False
'    sstab1.TabEnabled(2) = False
'    sstab1.TabEnabled(3) = False
'    Command3.Enabled = True
'    Command3.SetFocus
'End If
'End Sub

Private Sub S_ROU_LST_Click()
If CHECKTB("mar.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
End If

    Dim rs As New ADODB.Recordset
    If S_ROU_LST = "All" Then
        rs.Open "select * from mar", CON, adOpenKeyset
    Else
        rs.Open "select * from mar where rou_idy like '" & Trim(S_ROU_LST) & "'", CON, adOpenKeyset
    End If
    LIST3.CLEAR
    For A = 1 To rs.RecordCount
        LIST3.AddItem Trim(rs!mar_nme)
        rs.MoveNext
    Next
End Sub

Private Sub S_ROU_LST_KeyPress(K As Integer)
If K = 13 Then
    S_ROU_LST.Enabled = False
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(3) = False
    S_CMD.Enabled = True
    S_CMD.SetFocus
End If
End Sub


Private Sub smnid_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from smn where smn_idy like '" & Sincrement(smnid) & "'", CON, adOpenStatic
    If rs.RecordCount > 0 Then
        smnid = Sincrement(smnid)
        M_SMN_NME = Trim(rs("smn_nme"))
        smnid.Enabled = False
        Toolbar1.Buttons(12).Enabled = True
    Else
        If SSTab1.Tab = 0 Then
            list1.Tag = "Product"
        ElseIf SSTab1.Tab = 1 Then
            list1.Tag = "Brand"
        ElseIf SSTab1.Tab = 2 Then
            list1.Tag = "Segment"
        ElseIf SSTab1.Tab = 3 Then
            list1.Tag = "Category"
        End If
        smnlist
    End If
End If
End Sub

Private Sub smnlist()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    'cancel
    Exit Sub
End If

Dim ls As ListItem
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from smn where smn_nme like '" & smnid & "%'", CON, adOpenStatic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("smn_nme")))
        ls.ListSubItems.Add , , Trim(rs("smn_idy") & "")
        ls.ListSubItems.Add , , Trim(rs("smn_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 4
    list1.Visible = True
    list1.SetFocus
    'Toolbar1.Enabled = False
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        smnid.Enabled = True
        smnid.SetFocus
    End If
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

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select ID or Press <Esc> to Quit"
End Sub

Private Sub list1_KeyDown(KeyAscii As Integer, S As Integer)
Dim rs As New ADODB.Recordset
If KeyAscii = 13 Then
    If list1.ListItems.COUNT > 0 Then
        If rs.State = 1 Then rs.Close
        Dim lst As ListSubItem
        Set lst = list1.SelectedItem.ListSubItems.Item(1)
        rs.Open "select * from smn where smn_idy like '" & Trim(lst.Text) & "'", CON, adOpenStatic
        smnid = Trim(rs("smn_idy"))
        M_SMN_NME = Trim(rs!SMN_NME)
        list1.Visible = False
        If list1.Tag = "Product" Then
            SSTab1.Tab = 0
            Command3.Enabled = True
            Command3.SetFocus
        ElseIf list1.Tag = "Brand" Then
            SSTab1.Tab = 1
            B_CMD.Enabled = True
            B_CMD.SetFocus
        ElseIf list1.Tag = "Segment" Then
            SSTab1.Tab = 2
            S_CMD.Enabled = True
            S_CMD.SetFocus
        ElseIf list1.Tag = "Category" Then
            SSTab1.Tab = 3
            C_CMD.Enabled = True
            C_CMD.SetFocus
        End If
        
        smnid.Enabled = False
        MOD1 = True
        Toolbar1.Enabled = True
        Toolbar1.Buttons(12).Enabled = True
    End If
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 4 Then
    Frame9.Visible = False
    smnid.Visible = False
    Frame10.Visible = False
Else
    Frame9.Visible = True
    smnid.Visible = True
    Frame10.Visible = True
    If smnid.Enabled = True Then
        smnid.SetFocus
    Else
        If SSTab1.Tab = 0 Then
            MS.SetFocus
            MS.Row = 2
            MS.Col = 1
        ElseIf SSTab1.Tab = 1 Then
            msb.SetFocus
            msb.Row = 2
            msb.Col = 1
        ElseIf SSTab1.Tab = 2 Then
            MSS.SetFocus
            MSS.Row = 2
            MSS.Col = 1
        ElseIf SSTab1.Tab = 3 Then
            MSC.SetFocus
            MSC.Row = 2
            MSC.Col = 1
        End If
    End If
'    If ROU_LST.Enabled = True Then
'        ROU_LST.Selected(0) = True
'        ROU_LST.SetFocus
'    Else
'        MS.SetFocus
'    End If
'ElseIf SSTab1.Tab = 1 Then
'    If B_ROU_LST.Enabled = True Then
'        B_ROU_LST.Selected(0) = True
'        B_ROU_LST.SetFocus
'    Else
'        msb.SetFocus
'    End If
'ElseIf SSTab1.Tab = 2 Then
'    If S_ROU_LST.Enabled = True Then
'        S_ROU_LST.Selected(0) = True
'        S_ROU_LST.SetFocus
'    Else
'        MSS.SetFocus
'    End If
'ElseIf SSTab1.Tab = 3 Then
'    If C_ROU_LST.Enabled = True Then
'        C_ROU_LST.Selected(0) = True
'        C_ROU_LST.SetFocus
'    Else
'        msc.SetFocus
'    End If
End If
End Sub

Private Sub TGTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MS.Text = Int(Val(TGTQTY))
    If MS.Col <> 104 Then MS.Col = MS.Col + 1
    TGTQTY.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub TGTQTY_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "s" Then
    Cls1.TGT_SAVE
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "c" Then
    CANCELRECORD
ElseIf Button.KEY = "h" Then
M_HELP
End If
End Sub

Public Sub CANCELRECORD()
clrctr Me
M_SMN_NME = ""
MS.Rows = 2
MS.AddItem ""
msb.Rows = 2
msb.AddItem ""
MSS.Rows = 2
MSS.AddItem ""
MSC.Rows = 2
MSC.AddItem ""
Start = ""
ENDDATE = ""
B_START = ""
B_END = ""
C_START = ""
C_END = ""
S_START = ""
S_END = ""
M_BRD.Enabled = False
M_SEG.Enabled = False
M_CAT.Enabled = False
Command3.Enabled = False
B_CMD.Enabled = False
C_CMD.Enabled = False
S_CMD.Enabled = False
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = True
SSTab1.TabEnabled(2) = True
SSTab1.TabEnabled(3) = True
SSTab1.Tab = 0
'B_ROU_LST.Enabled = True
'S_ROU_LST.Enabled = True
'C_ROU_LST.Enabled = True
'ROU_LST.Enabled = True
smnid.Enabled = True
smnid.SetFocus
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.KEY = "SR" Then
    M_REP_TGT.Show 1
    CANCELRECORD
End If
End Sub
