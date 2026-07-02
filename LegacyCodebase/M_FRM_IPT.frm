VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_FRM_IPT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "External File Data Import"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "M_FRM_IPT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5150
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Import File Database"
      TabPicture(0)   =   "M_FRM_IPT.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cd1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtFileName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_IPT.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lv"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton Command3 
         Caption         =   "Clos&e"
         Height          =   375
         Left            =   7560
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   4680
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Default Values"
         Height          =   2535
         Left            =   2880
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   5655
         Begin VB.TextBox GRP_IDY 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1815
            MaxLength       =   10
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   240
            Width           =   1185
         End
         Begin VB.TextBox AREAID 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1815
            MaxLength       =   10
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   600
            Width           =   1185
         End
         Begin VB.TextBox marId 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1815
            MaxLength       =   10
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1185
         End
         Begin VB.TextBox Text32 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1815
            MaxLength       =   10
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1185
         End
         Begin VB.TextBox CATID 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1815
            MaxLength       =   10
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   960
            Width           =   1185
         End
         Begin VB.ComboBox ROU_LST 
            Height          =   315
            ItemData        =   "M_FRM_IPT.frx":047A
            Left            =   1815
            List            =   "M_FRM_IPT.frx":04E7
            Style           =   2  'Dropdown List
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2040
            Width           =   1800
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Group ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   0
            Left            =   735
            TabIndex        =   33
            Top             =   315
            Width           =   645
         End
         Begin VB.Label LABEL4 
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
            Left            =   3120
            TabIndex        =   32
            Top             =   240
            Width           =   2385
         End
         Begin VB.Label Label29 
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
            Left            =   3120
            TabIndex        =   28
            Top             =   600
            Width           =   2385
         End
         Begin VB.Label Label30 
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
            Left            =   3120
            TabIndex        =   27
            Top             =   1320
            Width           =   2385
         End
         Begin VB.Label Label19 
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
            Left            =   3120
            TabIndex        =   26
            Top             =   960
            Width           =   2385
         End
         Begin VB.Label Label31 
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
            Left            =   3120
            TabIndex        =   25
            Top             =   1680
            Width           =   2385
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Route (Beat)"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   24
            Top             =   2100
            Width           =   900
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prime Location ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   23
            Top             =   675
            Width           =   1260
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Market ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   4
            Left            =   675
            TabIndex        =   22
            Top             =   1395
            Width           =   705
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salesmen ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   5
            Left            =   480
            TabIndex        =   21
            Top             =   1755
            Width           =   900
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category ID"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   2
            Left            =   540
            TabIndex        =   20
            Top             =   1035
            Width           =   840
         End
      End
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2040
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   330
         Left            =   4005
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Select File Database"
         Enabled         =   0   'False
         Height          =   2535
         Left            =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   2295
         Begin VB.OptionButton opt_prd 
            Caption         =   "&Products (SKU)"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton opt_pgr 
            Caption         =   "SKU &Categories"
            Height          =   375
            Left            =   240
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1140
            Width           =   1815
         End
         Begin VB.OptionButton opt_act 
            Caption         =   "&Retailers"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Import"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7560
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   4200
         Width           =   975
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   720
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "xls"
         DialogTitle     =   "Select Excel File"
         FileName        =   "*.xls"
         Filter          =   "*.xls"
      End
      Begin MSComctlLib.ListView lv 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   7815
         _ExtentX        =   13785
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Frame Frame3 
         Caption         =   "Default Values"
         Height          =   2175
         Left            =   2880
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   5655
         Begin VB.TextBox grpid 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   360
            Width           =   1185
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1185
         End
         Begin VB.Label GRPSEG 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2595
            TabIndex        =   41
            Top             =   840
            Width           =   2955
            WordWrap        =   -1  'True
         End
         Begin VB.Label GRPCAT 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2595
            TabIndex        =   40
            Top             =   360
            Width           =   2955
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Group ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   39
            Top             =   435
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   300
            TabIndex        =   5
            Top             =   1395
            Width           =   780
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2595
            TabIndex        =   4
            Top             =   1320
            Width           =   2955
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   5280
         TabIndex        =   37
         Top             =   4320
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   3120
         TabIndex        =   36
         Top             =   4320
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records Imported"
         Height          =   195
         Left            =   3960
         TabIndex        =   35
         Top             =   4320
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Records"
         Height          =   195
         Left            =   1920
         TabIndex        =   34
         Top             =   4320
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label20 
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
         Left            =   -74640
         TabIndex        =   30
         Top             =   720
         Width           =   7815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select Worksheet"
         Height          =   195
         Left            =   600
         TabIndex        =   13
         Top             =   795
         Width           =   1275
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
            Picture         =   "M_FRM_IPT.frx":0554
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IPT.frx":0A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IPT.frx":0E60
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IPT.frx":11B4
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IPT.frx":17EC
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IPT.frx":1B40
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IPT.frx":1FA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IPT.frx":2294
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IPT.frx":26A0
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IPT.frx":27AC
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IPT.frx":2B00
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_IPT.frx":2F14
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
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
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "New (Ctrl+N)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "pre"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "d"
            Object.ToolTipText     =   "Delete (Ctrl+D)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_IPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsact As New ADODB.Recordset
Dim rsloc As New ADODB.Recordset
Dim rsmar As New ADODB.Recordset
Dim RSSMN As New ADODB.Recordset
Dim RSCAT As New ADODB.Recordset
Dim rspgr As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
Dim RSBRD As New ADODB.Recordset
Dim rsgrp As New ADODB.Recordset
Dim rsrdy As New ADODB.Recordset
Dim appXL As Excel.Application
Dim appX As Excel.Workbook
Dim R As Integer, C As Integer, A As Integer

Private Sub AREAID_keydown(K As Integer, Shift As Integer)
If CHECKTB("loc.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsloc As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If rsloc.State = 1 Then rsloc.Close
        rsloc.Open " select * from LOC where LOC_IDY like '" & Sincrement(AREAID) & "'", CON, adOpenStatic
        If rsloc.RecordCount > 0 Then
            AREAID = Sincrement(AREAID)
            Label29 = rsloc("LOC_NME")
            If K = 38 Then
                
            Else
                CATID.SetFocus
            End If
        Else
            arealist
        End If
End If
End Sub

Private Sub Command1_Click()
''ON ERROR GoTo ERRH
cd1.ShowOpen
cd1.CancelError = False
If Not Trim(cd1.FileName) = "" Then
    If Right(Trim(cd1.FileName), 3) = "xls" Then
        txtFileName.Text = cd1.FileName
        Set appX = GetObject(Trim(txtFileName.Text))
        Frame1.Enabled = True
        'opt_act.SetFocus
    Else
        MsgBox "Please select a Valid Excel File"
        txtFileName.Text = ""
        Exit Sub
    End If
End If
'C = 0
'If RSACT.State = 1 Then RSACT.Close
'RSACT.Open "SELECT * FROM ACT", CON, adOpenKeyset, adLockOptimistic
'RSACT.MoveLast
'For R = 2 To 100
'    If Trim(appXL.Sheets(1).Cells(R, 1)) = "" Then
'        Exit For
'    End If
'    If Trim(appXL.Sheets(1).Cells(R, 12)) = "" Then
'        MsgBox "Location ID Does Not Exist for " & R & " Record"
'        Exit Sub
'    ElseIf Trim(appXL.Sheets(1).Cells(R, 13)) = "" Then
'        MsgBox "Market ID Does Not Exist for " & R & " Record"
'        Exit Sub
'    ElseIf Trim(appXL.Sheets(1).Cells(R, 14)) = "" Then
'        MsgBox "Salesman ID Does Not Exist for " & R & " Record"
'        Exit Sub
'    ElseIf Trim(appXL.Sheets(1).Cells(R, 15)) = "" Then
'        MsgBox "Category ID Does Not Exist for " & R & " Record"
'        Exit Sub
'    End If
'Next
'For R = 2 To 100
'If RSLOC.State = 1 Then RSLOC.Close
'If RSMAR.State = 1 Then RSMAR.Close
'If RSSMN.State = 1 Then RSSMN.Close
'If RSCAT.State = 1 Then RSCAT.Close
'    If Trim(appXL.Sheets(1).Cells(R, 1)) = "" Then
'        Exit For
'    End If
'    RSACT.AddNew
'    For i = 0 To RSACT.Fields.Count - 1
'        If RSACT(i).Type = adVarChar Or RSACT(i).Type = adChar Or RSACT(i).Type = adBSTR Or RSACT(i).Type = adLongVarChar Or RSACT(i).Type = adLongVarWChar Then
'            RSACT(i) = ""
'        ElseIf RSACT(i).Type = adInteger Or RSACT(i).Type = adNumeric Or RSACT(i).Type = adBigInt Or RSACT(i).Type = adSingle Or RSACT(i).Type = adDouble Or RSACT(i).Type = adBigInt Or RSACT(i).Type = adCurrency Or RSACT(i).Type = adNumeric Or RSACT(i).Type = adSmallInt Or RSACT(i).Type = adDecimal Then
'            RSACT(i) = 0
'        ElseIf RSACT(i).Type = adDate Or RSACT(i).Type = adDBDate Or RSACT(i).Type = adDBTimeStamp Then
'            RSACT(i) = Date
'        End If
'    Next
'    RSACT("act_idy") = increment("IDY_ACT")
'    RSACT("grp_clf") = appXL.Sheets(1).Cells(R, 1)
'    If Trim(RSACT("GRP_CLF")) = "Debtor" Then
'        RSACT("grp_idy") = 7
'    ElseIf Trim(RSACT("grp_clf")) = "Creditor" Then
'        RSACT("grp_idy") = 8
'    End If
'    RSACT("act_nme") = appXL.Sheets(1).Cells(R, 2)
'    RSACT("act_gst") = appXL.Sheets(1).Cells(R, 3)
'    RSACT("act_cst") = appXL.Sheets(1).Cells(R, 4)
'    RSACT("act_ad1") = appXL.Sheets(1).Cells(R, 5)
'    RSACT("act_ad2") = appXL.Sheets(1).Cells(R, 6)
'    RSACT("act_ad3") = appXL.Sheets(1).Cells(R, 7)
'    RSACT("act_ad4") = appXL.Sheets(1).Cells(R, 8)
'    RSACT("act_zip") = appXL.Sheets(1).Cells(R, 9)
'    RSACT("act_tfx") = appXL.Sheets(1).Cells(R, 10)
'    RSACT("act_cpn") = appXL.Sheets(1).Cells(R, 11)
'    'RSACT("loc_idy") = appXL.Sheets(1).Cells(R, 12)
'    RSLOC.Open "SELECT * FROM LOC WHERE LOC_NME LIKE '" & Trim(appXL.Sheets(1).Cells(R, 12)) & "'", CON, adOpenKeyset, adLockOptimistic
'    If RSLOC.EOF = False Then
'        RSACT("LOC_IDY") = RSLOC("LOC_IDY")
'    Else
'        RSLOC.AddNew
'        RSLOC("LOC_IDY") = increment("IDY_LOC")
'        RSLOC("LOC_NME") = appXL.Sheets(1).Cells(R, 12)
'        RSLOC("LOC_TYP") = ""
'        RSACT("LOC_IDY") = RSLOC("LOC_IDY")
'        RSLOC.Update
'    End If
'    RSSMN.Open "SELECT * FROM SMN WHERE SMN_NME LIKE '" & Trim(appXL.Sheets(1).Cells(R, 14)) & "'", CON, adOpenKeyset, adLockOptimistic
'    If RSSMN.EOF = False Then
'        RSACT("SMN_IDY") = RSSMN("SMN_IDY")
'    Else
'        RSSMN.AddNew
'        For i = 0 To RSSMN.Fields.Count - 1
'        If RSSMN(i).Type = adVarChar Or RSSMN(i).Type = adChar Or RSSMN(i).Type = adBSTR Or RSSMN(i).Type = adLongVarChar Or RSSMN(i).Type = adLongVarWChar Then
'            RSSMN(i) = ""
'        ElseIf RSSMN(i).Type = adInteger Or RSSMN(i).Type = adNumeric Or RSSMN(i).Type = adBigInt Or RSSMN(i).Type = adSingle Or RSSMN(i).Type = adDouble Or RSSMN(i).Type = adBigInt Or RSSMN(i).Type = adCurrency Or RSSMN(i).Type = adNumeric Or RSSMN(i).Type = adSmallInt Or RSSMN(i).Type = adDecimal Then
'            RSSMN(i) = 0
'        ElseIf RSSMN(i).Type = adDate Or RSSMN(i).Type = adDBDate Or RSSMN(i).Type = adDBTimeStamp Then
'            RSSMN(i) = Date
'        End If
'        Next
'        RSSMN("SMN_IDY") = increment("IDY_SMN")
'        RSSMN("SMN_NME") = appXL.Sheets(1).Cells(R, 14)
'        RSACT("SMN_IDY") = RSSMN("SMN_IDY")
'        RSSMN.Update
'    End If
'    RSMAR.Open "SELECT * FROM MAR WHERE MAR_NME LIKE '" & Trim(appXL.Sheets(1).Cells(R, 13)) & "'", CON, adOpenKeyset, adLockOptimistic
'    If RSMAR.EOF = False Then
'        RSACT("MAR_IDY") = RSMAR("MAR_IDY")
'    Else
'        RSMAR.AddNew
'        For i = 0 To RSMAR.Fields.Count - 1
'        If RSMAR(i).Type = adVarChar Or RSMAR(i).Type = adChar Or RSMAR(i).Type = adBSTR Or RSMAR(i).Type = adLongVarChar Or RSMAR(i).Type = adLongVarWChar Then
'            RSMAR(i) = ""
'        ElseIf RSMAR(i).Type = adInteger Or RSMAR(i).Type = adNumeric Or RSMAR(i).Type = adBigInt Or RSMAR(i).Type = adSingle Or RSMAR(i).Type = adDouble Or RSMAR(i).Type = adBigInt Or RSMAR(i).Type = adCurrency Or RSMAR(i).Type = adNumeric Or RSMAR(i).Type = adSmallInt Or RSMAR(i).Type = adDecimal Then
'            RSMAR(i) = 0
'        ElseIf RSMAR(i).Type = adDate Or RSMAR(i).Type = adDBDate Or RSMAR(i).Type = adDBTimeStamp Then
'            RSMAR(i) = Date
'        End If
'        Next
'        RSMAR("MAR_IDY") = increment("IDY_MAR")
'        RSMAR("MAR_NME") = appXL.Sheets(1).Cells(R, 13)
'        RSMAR("SMN_IDY") = RSACT("SMN_IDY")
'        RSACT("MAR_IDY") = RSMAR("MAR_IDY")
'        RSMAR.Update
'    End If
'    'RSACT("Mar_idy") = appXL.Sheets(1).Cells(R, 13)
'    'RSACT("smn_Idy") = appXL.Sheets(1).Cells(R, 14)
'    'RSACT("Cat_idy") = appXL.Sheets(1).Cells(R, 15)
'    RSCAT.Open "SELECT * FROM CAT WHERE CAT_NME LIKE '" & Trim(appXL.Sheets(1).Cells(R, 15)) & "'", CON, adOpenKeyset, adLockOptimistic
'    If RSCAT.EOF = False Then
'        RSACT("CAT_IDY") = RSCAT("CAT_IDY")
'    Else
'        RSCAT.AddNew
'        For i = 0 To RSCAT.Fields.Count - 1
'        If RSCAT(i).Type = adVarChar Or RSCAT(i).Type = adChar Or RSCAT(i).Type = adBSTR Or RSCAT(i).Type = adLongVarChar Or RSCAT(i).Type = adLongVarWChar Then
'            RSCAT(i) = ""
'        ElseIf RSCAT(i).Type = adInteger Or RSCAT(i).Type = adNumeric Or RSCAT(i).Type = adBigInt Or RSCAT(i).Type = adSingle Or RSCAT(i).Type = adDouble Or RSCAT(i).Type = adBigInt Or RSCAT(i).Type = adCurrency Or RSCAT(i).Type = adNumeric Or RSCAT(i).Type = adSmallInt Or RSCAT(i).Type = adDecimal Then
'            RSCAT(i) = 0
'        ElseIf RSCAT(i).Type = adDate Or RSCAT(i).Type = adDBDate Or RSCAT(i).Type = adDBTimeStamp Then
'            RSCAT(i) = Date
'        End If
'        Next
'        RSCAT("CAT_IDY") = increment("IDY_CAT")
'        RSCAT("CAT_NME") = appXL.Sheets(1).Cells(R, 15)
'        RSACT("CAT_IDY") = RSCAT("CAT_IDY")
'        RSCAT.Update
'    End If
'    RSACT("rou_idy") = appXL.Sheets(1).Cells(R, 16)
'    RSACT("dbt_typ") = "Y"
'    RSACT("key_act") = "N"
'    RSACT("dbt_rap") = "B"
'    RSACT("act_typ") = "Y"
'    RSACT.Update
'    C = C + 1
'Next
'MsgBox "Export Completed"
Exit Sub
ERRH:
MsgBox "Select a Valid file Database", vbInformation, "Effmcg"
txtFileName.SetFocus
SendKeys "{HOME}+{END}"
End Sub

Private Sub Command2_Click()
'If appXL.State = 1 Then appXL.Close
''ON ERROR GoTo ERRH1
    Set appX = GetObject(Trim(txtFileName.Text))
    If opt_act.Value = True Then
        act
    ElseIf opt_pgr.Value = True Then
        pgr
    ElseIf opt_prd.Value = True Then
        prd
    End If
'ERRH1:
'MsgBox "Select a File", vbExclamation, "Effmcg"
'txtFileName.SetFocus
'SendKeys "{home}+{end}"
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Activate()
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
txtFileName.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If K = 120 Then Shell "c:\windows\calc.exe"
If K = 81 And S = 2 Then
    CANCELRECORD
ElseIf K = 88 And S = 2 Then
    Unload Me
ElseIf K = 112 Then
    M_HELP
ElseIf K = 80 And S = 2 Then
    M_FRM_AQY.Tag = "ACT"
    M_FRM_AQY.Show 1
ElseIf K = 82 And S = 2 Then
    'MDIForm1.RRReport1.ReportName = App.Path & "\Accounts list.RRW"
    'MDIForm1.RRReport1.Destination = 1
    'MDIForm1.RRReport1.Printer = "?"
    'MDIForm1.RRReport1.RunReport 1
End If
End Sub

Public Sub act()
If CHECKTB("act.dbf,rdy.dbf,loc.dbf,smn.dbf,mar.dbf,cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
If Trim(GRP_IDY) = "" Then
    MsgBox "Please select a Group A/C", vbInformation, "Effmcg"
    GRP_IDY.SetFocus
    Exit Sub
End If
If MsgBox("Import from " & txtFileName, vbQuestion + vbYesNo) = vbNo Then Exit Sub
Command2.Enabled = False
If rsgrp.State = 1 Then rsgrp.Close
rsgrp.Open "select * from grp where grp_IDY like '" & Trim(GRP_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
C = 0
If UCase(Trim(appX.Sheets(1).Cells(1, 1))) = "NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 2))) = "GST" And UCase(Trim(appX.Sheets(1).Cells(1, 3))) = "CST" And UCase(Trim(appX.Sheets(1).Cells(1, 4))) = "ADDRESS1" And UCase(Trim(appX.Sheets(1).Cells(1, 5))) = "ADDRESS2" And UCase(Trim(appX.Sheets(1).Cells(1, 6))) = "ADDRESS3" And UCase(Trim(appX.Sheets(1).Cells(1, 7))) = "ADDRESS4" And UCase(Trim(appX.Sheets(1).Cells(1, 8))) = "ZIPCODE" And UCase(Trim(appX.Sheets(1).Cells(1, 9))) = "TELEPHONE" And UCase(Trim(appX.Sheets(1).Cells(1, 10))) = "CONTACT" And UCase(Trim(appX.Sheets(1).Cells(1, 11))) = "LOCATION" And UCase(Trim(appX.Sheets(1).Cells(1, 12))) = "MARKET" And UCase(Trim(appX.Sheets(1).Cells(1, 13))) = "SALESMAN" And UCase(Trim(appX.Sheets(1).Cells(1, 14))) = "CATEGORY" And UCase(Trim(appX.Sheets(1).Cells(1, 15))) = "ROUTE" Then
Me.MousePointer = vbHourglass
    If rsact.State = 1 Then rsact.Close
    'If rsrdy.State = 1 Then rsrdy.Close
    rsact.Open "SELECT * FROM ACT", CON, adOpenKeyset, adLockOptimistic
    rsrdy.Open "select * from rdy where act_idy=''", CON, adOpenKeyset, adLockOptimistic
'    rsact.MoveLast
'    For R = 2 To 100
'        If Trim(appXL.Sheets(1).Cells(R, 1)) = "" Then
'            Exit For
'        End If
'        If Trim(appXL.Sheets(1).Cells(R, 11)) = "" Then
'            MsgBox "Location ID Does Not Exist for " & R - 1 & " Record"
'            'Set appXL = Nothing
'            Exit Sub
'        ElseIf Trim(appXL.Sheets(1).Cells(R, 12)) = "" Then
'            MsgBox "Market ID Does Not Exist for " & R - 1 & " Record"
'            'Set appXL = Nothing
'            Exit Sub
'        ElseIf Trim(appXL.Sheets(1).Cells(R, 13)) = "" Then
'            MsgBox "Salesman ID Does Not Exist for " & R - 1 & " Record"
'            'Set appXL = Nothing
'            Exit Sub
'        ElseIf Trim(appXL.Sheets(1).Cells(R, 14)) = "" Then
'            MsgBox "Category ID Does Not Exist for " & R - 1 & " Record"
'            'Set appXL = Nothing
'            Exit Sub
'        End If
'    Next
    appX.Sheets(1).Cells(1, 27).Formula = "=counta(a1:a65536)"
    maxrows = appX.Sheets(1).Cells(1, 27)
    Label7 = maxrows - 1
    For R = 2 To maxrows + 1
    SHOWBUT (True)
    If rsloc.State = 1 Then rsloc.Close
    If rsmar.State = 1 Then rsmar.Close
    If RSSMN.State = 1 Then RSSMN.Close
    If RSCAT.State = 1 Then RSCAT.Close
        If Trim(appX.Sheets(1).Cells(R, 1)) = "" Then
            Exit For
        End If
        rsact.AddNew
        For I = 0 To rsact.Fields.Count - 1
            If rsact(I).Type = adVarChar Or rsact(I).Type = adChar Or rsact(I).Type = adBSTR Or rsact(I).Type = adLongVarChar Or rsact(I).Type = adLongVarWChar Then
                rsact(I) = " "
            ElseIf rsact(I).Type = adInteger Or rsact(I).Type = adNumeric Or rsact(I).Type = adBigInt Or rsact(I).Type = adSingle Or rsact(I).Type = adDouble Or rsact(I).Type = adBigInt Or rsact(I).Type = adCurrency Or rsact(I).Type = adNumeric Or rsact(I).Type = adSmallInt Or rsact(I).Type = adDecimal Then
                rsact(I) = 0
            ElseIf rsact(I).Type = adDate Or rsact(I).Type = adDBDate Or rsact(I).Type = adDBTimeStamp Then
                rsact(I) = Date
            End If
        Next
        rsact("act_idy") = increment("IDY_ACT")
        rsact("grp_clf") = rsgrp("GRP_CLF")
'       If Trim(RSACT("GRP_CLF")) = "Debtor" Then
        rsact("grp_idy") = rsgrp("GRP_IDY")
'       ElseIf Trim(RSACT("grp_clf")) = "Creditor" Then
'           RSACT("grp_idy") = "0000000008"
'       End If
        rsact("act_nme") = ADS(rsact, "ACT_NME", appX.Sheets(1).Cells(R, 1))
        rsact("act_gst") = ADS(rsact, "act_gst", appX.Sheets(1).Cells(R, 2))
        rsact("act_cst") = ADS(rsact, "act_cst", appX.Sheets(1).Cells(R, 3))
        rsact("act_ad1") = ADS(rsact, "act_ad1", appX.Sheets(1).Cells(R, 4))
        rsact("act_ad2") = ADS(rsact, "act_ad2", appX.Sheets(1).Cells(R, 5))
        rsact("act_ad3") = ADS(rsact, "act_ad3", appX.Sheets(1).Cells(R, 6))
        rsact("act_ad4") = ADS(rsact, "act_ad4", appX.Sheets(1).Cells(R, 7))
        rsact("act_zip") = ADS(rsact, "act_zip", appX.Sheets(1).Cells(R, 8))
        rsact("act_tfx") = ADS(rsact, "act_tfx", appX.Sheets(1).Cells(R, 9))
        rsact("act_cpn") = ADS(rsact, "act_cpn", appX.Sheets(1).Cells(R, 10))
        rsact("cst_typ") = "Retailer"
        If Trim(appX.Sheets(1).Cells(R, 11)) <> "" Then
            If rsloc.State = 1 Then rsloc.Close
            rsloc.Open "SELECT * FROM LOC WHERE LOC_NME LIKE '" & Trim(appX.Sheets(1).Cells(R, 11)) & "'", CON, adOpenKeyset, adLockOptimistic
            If rsloc.EOF = False Then
                rsact("LOC_IDY") = rsloc("LOC_IDY")
            Else
                rsloc.AddNew
                rsloc("LOC_IDY") = increment("IDY_LOC")
                rsloc("LOC_NME") = ADS(rsloc, "loc_nme", appX.Sheets(1).Cells(R, 11))
                rsloc("LOC_TYP") = ""
                rsact("LOC_IDY") = rsloc("LOC_IDY")
                rsloc.Update
            End If
        Else
            rsact("LOC_IDY") = Trim(AREAID)
        End If
        If Trim(appX.Sheets(1).Cells(R, 13)) <> "" Then
            If RSSMN.State = 1 Then RSSMN.Close
            RSSMN.Open "SELECT * FROM SMN WHERE SMN_NME LIKE '" & Trim(appX.Sheets(1).Cells(R, 13)) & "'", CON, adOpenKeyset, adLockOptimistic
            If RSSMN.EOF = False Then
                rsact("SMN_IDY") = RSSMN("SMN_IDY")
            Else
                RSSMN.AddNew
                For I = 0 To RSSMN.Fields.Count - 1
                If RSSMN(I).Type = adVarChar Or RSSMN(I).Type = adChar Or RSSMN(I).Type = adBSTR Or RSSMN(I).Type = adLongVarChar Or RSSMN(I).Type = adLongVarWChar Then
                    RSSMN(I) = ""
                ElseIf RSSMN(I).Type = adInteger Or RSSMN(I).Type = adNumeric Or RSSMN(I).Type = adBigInt Or RSSMN(I).Type = adSingle Or RSSMN(I).Type = adDouble Or RSSMN(I).Type = adBigInt Or RSSMN(I).Type = adCurrency Or RSSMN(I).Type = adNumeric Or RSSMN(I).Type = adSmallInt Or RSSMN(I).Type = adDecimal Then
                    RSSMN(I) = 0
                ElseIf RSSMN(I).Type = adDate Or RSSMN(I).Type = adDBDate Or RSSMN(I).Type = adDBTimeStamp Then
                    RSSMN(I) = Date
                End If
                Next
                RSSMN("SMN_IDY") = increment("IDY_SMN")
                RSSMN("SMN_NME") = ADS(RSSMN, "smn_nme", appX.Sheets(1).Cells(R, 13))
                rsact("SMN_IDY") = RSSMN("SMN_IDY")
                RSSMN.Update
            End If
        Else
            rsact("SMN_IDY") = Trim(Text32)
        End If
        If Trim(appX.Sheets(1).Cells(R, 12)) <> "" Then
            If rsmar.State = 1 Then rsmar.Close
            rsmar.Open "SELECT * FROM MAR WHERE MAR_NME LIKE '" & Trim(appX.Sheets(1).Cells(R, 12)) & "'", CON, adOpenKeyset, adLockOptimistic
            If rsmar.EOF = False Then
                rsact("MAR_IDY") = rsmar("MAR_IDY")
            Else
                rsmar.AddNew
                For I = 0 To rsmar.Fields.Count - 1
                If rsmar(I).Type = adVarChar Or rsmar(I).Type = adChar Or rsmar(I).Type = adBSTR Or rsmar(I).Type = adLongVarChar Or rsmar(I).Type = adLongVarWChar Then
                    rsmar(I) = ""
                ElseIf rsmar(I).Type = adInteger Or rsmar(I).Type = adNumeric Or rsmar(I).Type = adBigInt Or rsmar(I).Type = adSingle Or rsmar(I).Type = adDouble Or rsmar(I).Type = adBigInt Or rsmar(I).Type = adCurrency Or rsmar(I).Type = adNumeric Or rsmar(I).Type = adSmallInt Or rsmar(I).Type = adDecimal Then
                    rsmar(I) = 0
                ElseIf rsmar(I).Type = adDate Or rsmar(I).Type = adDBDate Or rsmar(I).Type = adDBTimeStamp Then
                    rsmar(I) = Date
                End If
                Next
                rsmar("MAR_IDY") = increment("IDY_MAR")
                rsmar("MAR_NME") = ADS(rsmar, "mar_nme", appX.Sheets(1).Cells(R, 12))
                rsmar("SMN_IDY") = rsact("SMN_IDY")
                rsmar("ROU_IDY") = "B"
                rsact("MAR_IDY") = rsmar("MAR_IDY")
                rsmar.Update
            End If
        Else
            rsact("mar_idy") = Trim(marId)
        End If
        'RSACT("Mar_idy") = appX.Sheets(1).Cells(R, 13)
        'RSACT("smn_Idy") = appX.Sheets(1).Cells(R, 14)
        'RSACT("Cat_idy") = appX.Sheets(1).Cells(R, 15)
        If Trim(appX.Sheets(1).Cells(R, 14)) <> "" Then
            If RSCAT.State = 1 Then RSCAT.Close
            RSCAT.Open "SELECT * FROM CAT WHERE CAT_NME LIKE '" & Trim(appX.Sheets(1).Cells(R, 14)) & "'", CON, adOpenKeyset, adLockOptimistic
            If RSCAT.EOF = False Then
                rsact("CAT_IDY") = RSCAT("CAT_IDY")
            Else
                RSCAT.AddNew
                For I = 0 To RSCAT.Fields.Count - 1
                If RSCAT(I).Type = adVarChar Or RSCAT(I).Type = adChar Or RSCAT(I).Type = adBSTR Or RSCAT(I).Type = adLongVarChar Or RSCAT(I).Type = adLongVarWChar Then
                    RSCAT(I) = " "
                ElseIf RSCAT(I).Type = adInteger Or RSCAT(I).Type = adNumeric Or RSCAT(I).Type = adBigInt Or RSCAT(I).Type = adSingle Or RSCAT(I).Type = adDouble Or RSCAT(I).Type = adBigInt Or RSCAT(I).Type = adCurrency Or RSCAT(I).Type = adNumeric Or RSCAT(I).Type = adSmallInt Or RSCAT(I).Type = adDecimal Then
                    RSCAT(I) = 0
                ElseIf RSCAT(I).Type = adDate Or RSCAT(I).Type = adDBDate Or RSCAT(I).Type = adDBTimeStamp Then
                    RSCAT(I) = Date
                End If
                Next
                RSCAT("CAT_IDY") = increment("IDY_CAT")
                RSCAT("CAT_NME") = ADS(RSCAT, "cat_nme", appX.Sheets(1).Cells(R, 14))
                rsact("CAT_IDY") = RSCAT("CAT_IDY")
                RSCAT.Update
            End If
        Else
            rsact("cat_idy") = Trim(CATID)
        End If
        If UCase(appX.Sheets(1).Cells(R, 15)) <> "" Then
            rsact("rou_idy") = ADS(rsact, "rou_idy", UCase(appX.Sheets(1).Cells(R, 15)))
        Else
            rsact("rou_idy") = Trim(ROU_LST)
        End If
        rsact("dbt_typ") = "Y"
        rsact("key_act") = "N"
        rsact("dbt_rap") = "B"
        rsact("act_typ") = "SAL"
        rsact.Update
        For d = 1 To 7
            rsrdy.AddNew
            For K = 0 To rsrdy.Fields.Count - 1
                If rsrdy(K).Type = adVarChar Or rsrdy(K).Type = adChar Or rsrdy(K).Type = adBSTR Or rsrdy(K).Type = adLongVarChar Or rsrdy(K).Type = adLongVarWChar Then
                    rsrdy(K) = " "
                ElseIf rsrdy(K).Type = adInteger Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adSingle Or rsrdy(K).Type = adDouble Or rsrdy(K).Type = adBigInt Or rsrdy(K).Type = adCurrency Or rsrdy(K).Type = adNumeric Or rsrdy(K).Type = adSmallInt Or rsrdy(K).Type = adDecimal Then
                    rsrdy(K) = 0
                ElseIf rsrdy(K).Type = adDate Or rsrdy(K).Type = adDBDate Or rsrdy(K).Type = adDBTimeStamp Then
                    rsrdy(K) = Date
                End If
            Next
            rsrdy("act_idy") = rsact("act_idy")
            rsrdy("day_idy") = d
            rsrdy("smn_idy") = rsact("smn_idy")
            rsrdy("rou_idy") = rsact("rou_idy")
        Next
        C = C + 1
        Label8 = R - 1
        Label8.Refresh
    Next
    Me.MousePointer = vbDefault
    MsgBox "Import Completed", vbExclamation
    appX.Close 0
    appXL.Quit
    Set appXL = Nothing
    Unload Me
Else
    MsgBox "Invalid File Selected for Accounts & Parties Import", vbExclamation
End If
End Sub

Public Sub pgr()
If MsgBox("Import from " & txtFileName, vbQuestion + vbYesNo) = vbNo Then Exit Sub

If UCase(Trim(appX.Sheets(1).Cells(1, 1))) = "CATEGORIES" And UCase(Trim(appX.Sheets(1).Cells(1, 2))) = "SEGMENT" Then
    Me.MousePointer = vbHourglass
    If rspgr.State = 1 Then rspgr.Close
    rspgr.Open "select * from pgr", CON, adOpenKeyset, adLockOptimistic
    appX.Sheets(1).Cells(1, 27).Formula = "=counta(a1:a65536)"
    maxrows = appX.Sheets(1).Cells(1, 27)
    Label7 = maxrows - 1
    SHOWBUT (True)
    For R = 2 To maxrows + 1
        If Trim(appX.Sheets(1).Cells(R, 1)) = "" Then
            Exit For
        End If
        rspgr.AddNew
        rspgr("Grp_idy") = increment("idy_Pgr")
        rspgr("grp_cat") = ADS(rspgr, "grp_cat", Trim(appX.Sheets(1).Cells(R, 1)))
        rspgr("grp_cls") = ADS(rspgr, "grp_cls", Trim(appX.Sheets(1).Cells(R, 2)))
        rspgr.Update
        Label8 = R - 1
        Label8.Refresh
    Next
    Me.MousePointer = vbDefault
    MsgBox "Import Completed", vbExclamation
    appX.Close 0
    appXL.Quit
    Set appXL = Nothing
    Unload Me
Else
    MsgBox "Invalid File Selected for Product Categories Import", vbExclamation
End If
End Sub
Public Sub prd()
If MsgBox("Import from " & txtFileName, vbQuestion + vbYesNo) = vbNo Then Exit Sub
If Trim(appX.Sheets(1).Cells(R, 6)) = "" And Trim(appX.Sheets(1).Cells(R, 7)) = "" Then
    If Trim(grpid) = "" Then
        MsgBox "Please select Product Group", vbExclamation
        grpid.SetFocus
        Exit Sub
    End If
End If
If rspgr.State = 1 Then rspgr.Close
If rsact.State = 1 Then rsact.Close
'A = Split(cmbSegment.Text, "-", , vbTextCompare)
'rspgr.Open "SELECT * FROM PGR WHERE GRP_CAT LIKE '" & Trim(Text2) & "' AND GRP_CLS LIKE '" & Label11 & "'", CON, adOpenKeyset, adLockOptimistic
'RSACT.Open "SELECT * FROM ACT WHERE ACT_idy LIKE '" & Trim(Text3) & "'", CON, adOpenKeyset, adLockOptimistic
If UCase(Trim(appX.Sheets(1).Cells(1, 1))) = "SKU NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 2))) = "COMPANY ID" And UCase(Trim(appX.Sheets(1).Cells(1, 3))) = "SHORT NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 4))) = "BRAND NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 5))) = "GROUP ID" And UCase(Trim(appX.Sheets(1).Cells(1, 6))) = "CATEGORY" And UCase(Trim(appX.Sheets(1).Cells(1, 7))) = "SEGMENT" And UCase(Trim(appX.Sheets(1).Cells(1, 8))) = "COMPANY NAME" And UCase(Trim(appX.Sheets(1).Cells(1, 9))) = "PUR RATE" And UCase(Trim(appX.Sheets(1).Cells(1, 10))) = "SPL RATE" And UCase(Trim(appX.Sheets(1).Cells(1, 11))) = "SAL RATE" And UCase(Trim(appX.Sheets(1).Cells(1, 12))) = "MRP RATE" And UCase(Trim(appX.Sheets(1).Cells(1, 13))) = "CON FACTOR" And UCase(Trim(appX.Sheets(1).Cells(1, 14))) = "SUB UNITS" Then
Me.MousePointer = vbHourglass
    If rsprd.State = 1 Then rsprd.Close
    If RSBRD.State = 1 Then RSBRD.Close
    rsprd.Open "select * from Prd", CON, adOpenKeyset, adLockOptimistic
    RSBRD.Open "select * from Brd", CON, adOpenKeyset, adLockOptimistic
    appX.Sheets(1).Cells(1, 27).Formula = "=counta(a1:a65536)"
    maxrows = appX.Sheets(1).Cells(1, 27)
    Label7 = maxrows - 1
    SHOWBUT (True)
    For R = 2 To maxrows + 1
        If Trim(appX.Sheets(1).Cells(R, 1)) = "" Then
            Exit For
        End If
        rsprd.AddNew
        For I = 0 To rsprd.Fields.Count - 1
            If rsprd(I).Type = adVarChar Or rsprd(I).Type = adChar Or rsprd(I).Type = adBSTR Or rsprd(I).Type = adLongVarChar Or rsprd(I).Type = adLongVarWChar Then
                rsprd(I) = ""
            ElseIf rsprd(I).Type = adInteger Or rsprd(I).Type = adNumeric Or rsprd(I).Type = adBigInt Or rsprd(I).Type = adSingle Or rsprd(I).Type = adDouble Or rsprd(I).Type = adBigInt Or rsprd(I).Type = adCurrency Or rsprd(I).Type = adNumeric Or rsprd(I).Type = adSmallInt Or rsprd(I).Type = adDecimal Then
                rsprd(I) = 0
            ElseIf rsprd(I).Type = adDate Or rsprd(I).Type = adDBDate Or rsprd(I).Type = adDBTimeStamp Then
                rsprd(I) = Date
            End If
        Next
        rsprd("prd_idy") = increment("idy_Prd")
        rsprd("prd_nme") = ADS(rsprd, "prd_nme", Trim(appX.Sheets(1).Cells(R, 1)))
        rsprd("alt_idy") = ADS(rsprd, "alt_idy", Trim(appX.Sheets(1).Cells(R, 2)))
        rsprd("stk_typ") = "Salable"
        rsprd("prd_prn") = ADS(rsprd, "prd_prn", Trim(appX.Sheets(1).Cells(R, 3)))
        rsprd("brd_nme") = ADS(rsprd, "brd_nme", Trim(appX.Sheets(1).Cells(R, 4)))
'       RSPRD("grp_idy") = rspgr("GRP_IDY") 'Trim(appX.Sheets(1).Cells(R, 6))
        rsprd("cmp_idy") = Trim(Text3) & "" 'RSACT("ACT_IDY") 'Trim(appX.Sheets(1).Cells(R, 7))
        rsprd("pur_rat") = ADS(rsprd, "pur_rat", convert(AMT(Trim(appX.Sheets(1).Cells(R, 9)))))
        rsprd("bas_rat") = ADS(rsprd, "bas_rat", convert(AMT(Trim(appX.Sheets(1).Cells(R, 10)))))
        rsprd("sal_rat") = ADS(rsprd, "sal_rat", convert(AMT(Trim(appX.Sheets(1).Cells(R, 11)))))
        rsprd("mrp_rat") = ADS(rsprd, "mrp_rat", convert(AMT(Trim(appX.Sheets(1).Cells(R, 12)))))
        rsprd("con_fac") = ADS(rsprd, "con_fac", Trim(appX.Sheets(1).Cells(R, 13)))
        'RSPRD("prd_uom") = Trim(appX.Sheets(1).Cells(R, 18))
        rsprd("sub_unt") = ADS(rsprd, "sub_unt", Trim(appX.Sheets(1).Cells(R, 14)))
        If Trim(appX.Sheets(1).Cells(R, 6)) <> "" And Trim(appX.Sheets(1).Cells(R, 7)) <> "" Then
            If rspgr.State = 1 Then rspgr.Close
            rspgr.Open "SELECT * FROM PGR WHERE grp_cat ='" & Trim(appX.Sheets(1).Cells(R, 6)) & "' and grp_cls ='" & Trim(appX.Sheets(1).Cells(R, 7)) <> "" & "'", CON, adOpenKeyset, adLockOptimistic
            If rspgr.EOF = False Then
                rsprd("grp_IDY") = rspgr("grp_IDY")
            Else
                rspgr.AddNew
                For I = 0 To rspgr.Fields.Count - 1
                If rspgr(I).Type = adVarChar Or rspgr(I).Type = adChar Or rspgr(I).Type = adBSTR Or rspgr(I).Type = adLongVarChar Or rspgr(I).Type = adLongVarWChar Then
                    rspgr(I) = ""
                ElseIf rspgr(I).Type = adInteger Or rspgr(I).Type = adNumeric Or rspgr(I).Type = adBigInt Or rspgr(I).Type = adSingle Or rspgr(I).Type = adDouble Or rspgr(I).Type = adBigInt Or rspgr(I).Type = adCurrency Or rspgr(I).Type = adNumeric Or rspgr(I).Type = adSmallInt Or rspgr(I).Type = adDecimal Then
                    rspgr(I) = 0
                ElseIf rspgr(I).Type = adDate Or rspgr(I).Type = adDBDate Or rspgr(I).Type = adDBTimeStamp Then
                    rspgr(I) = Date
                End If
                Next
                rspgr("grp_IDY") = increment("IDY_pgR")
                rspgr("grp_cat") = ADS(rspgr, "grp_cat", Trim(appX.Sheets(1).Cells(R, 6)))
                rspgr("grp_cls") = ADS(rspgr, "grp_cls", Trim(appX.Sheets(1).Cells(R, 7)))
                rsprd("grp_IDY") = rspgr("grp_IDY")
                rspgr.Update
            End If
        Else
            rsprd("grp_idy") = Trim(grpid)
        End If
        rsprd.Update
        RSBRD.AddNew
        For I = 0 To RSBRD.Fields.Count - 1
            If RSBRD(I).Type = adVarChar Or RSBRD(I).Type = adChar Or RSBRD(I).Type = adBSTR Or RSBRD(I).Type = adLongVarChar Or RSBRD(I).Type = adLongVarWChar Then
                RSBRD(I) = ""
            ElseIf RSBRD(I).Type = adInteger Or RSBRD(I).Type = adNumeric Or RSBRD(I).Type = adBigInt Or RSBRD(I).Type = adSingle Or RSBRD(I).Type = adDouble Or RSBRD(I).Type = adBigInt Or RSBRD(I).Type = adCurrency Or RSBRD(I).Type = adNumeric Or RSBRD(I).Type = adSmallInt Or RSBRD(I).Type = adDecimal Then
                RSBRD(I) = 0
            ElseIf RSBRD(I).Type = adDate Or RSBRD(I).Type = adDBDate Or RSBRD(I).Type = adDBTimeStamp Then
                RSBRD(I) = Date
            End If
        Next
        RSBRD("bat_idy") = "BAT" & T7increment("idy_bat")
        RSBRD("prd_idy") = rsprd("PRD_IDY")
        RSBRD("PRD_PDR") = rsprd("PUR_RAT")
        RSBRD("PRD_SDR") = rsprd("SAL_RAT")
        RSBRD("PRD_MRP") = rsprd("MRP_RAT")
        RSBRD("BRD_DOM") = Date
        RSBRD("BRD_DOE") = Date
        RSBRD("bat_nme") = " "
        RSBRD.Update
        Label8 = R - 1
        Label8.Refresh
    Next
    Me.MousePointer = vbDefault
    MsgBox "Import Completed", vbExclamation
    appX.Close 0
    appXL.Quit
    Set appXL = Nothing
    Unload Me
Else
    MsgBox "Invalid File Selected for Product Data Import", vbExclamation
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub



Private Sub Form_Load()
Set appXL = CreateObject("Excel.Application")
Set appX = appXL.Workbooks.Add
End Sub

Private Sub Form_Terminate()
rsact.Close
rsloc.Close
rsmar.Close
RSSMN.Close
RSCAT.Close
rspgr.Close
rsprd.Close
RSBRD.Close
rsgrp.Close
rsrdy.Close
Set rsact = Nothing
Set rsloc = Nothing
Set rsmar = Nothing
Set RSSMN = Nothing
Set RSCAT = Nothing
Set rspgr = Nothing
Set rsprd = Nothing
Set RSBRD = Nothing
Set rsgrp = Nothing
Set rsrdy = Nothing
End Sub

Private Sub GRP_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Group Account ID or Press <Enter> to get List of Group ID's"
End Sub


Private Sub GRP_IDY_KeyDown(K As Integer, S As Integer)
If CHECKTB("grp.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
If K = 13 Or K = 40 Then
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
    RS.Open "select * from grp where grp_idy like '" & Sincrement(Trim(GRP_IDY)) & "'", CON, adOpenKeyset
    If RS.RecordCount > 0 Then
        GRP_IDY = Sincrement(GRP_IDY)
        Label4 = Trim(RS("grp_nme"))
        AREAID.SetFocus
    Else
        grouplist
    End If
End If
End Sub

Private Sub grouplist()
If CHECKTB("grp.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    
    Exit Sub
End If
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from grp where grp_nme like '" & Trim(GRP_IDY) & "%'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    lv.ListItems.Clear
    For A = 1 To RS.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(RS("grp_nme")))
        ls.ListSubItems.Add , , Trim(RS("grp_idy"))
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Group Account"
    lv.Visible = True
    lv.SetFocus
Else
   MsgBox "Group Account Not Found, Define Group Account !", vbExclamation, "Effmcg"
   SSTab1.Tab = 0
   GRP_IDY.SetFocus
End If
End Sub

Private Sub GRP_IDY_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub grpid_KeyDown(KeyCode As Integer, Shift As Integer)
If CHECKTB("pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim RS As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
    If RS.State = 1 Then RS.Close
        RS.Open "select * from pgr where grp_idy like '" & Sincrement(grpid) & "'", CON, adOpenStatic, adLockReadOnly
        If RS.RecordCount > 0 Then
        grpid = Sincrement(Trim(Text1))
        GRPCAT = Trim(RS("grp_cat"))
        GRPSEG = Trim(RS("grp_cls"))
        grpid.Enabled = False
        Text3.SetFocus
        MOD1 = True
        'Toolbar1.Buttons(12).Enabled = True
        Else
            prdgrplist
        End If
End If
End Sub

Private Sub opt_act_Click()
opt_act.Enabled = True
opt_prd.Enabled = False
opt_pgr.Enabled = False
'Frame1.Enabled = False
Frame3.Visible = False
Frame2.Visible = True
GRP_IDY.SetFocus
Command2.Enabled = True
End Sub

Private Sub opt_pgr_Click()
opt_act.Enabled = False
opt_prd.Enabled = False
opt_pgr.Enabled = True
Frame2.Visible = False
Frame3.Visible = False
Command2.Enabled = True
End Sub

Private Sub opt_pgr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command2.SetFocus
End If
End Sub

Private Sub opt_prd_Click()
opt_act.Enabled = False
opt_prd.Enabled = True
opt_pgr.Enabled = False
'Frame1.Enabled = False
Frame2.Visible = False
Frame3.Visible = True
grpid.SetFocus
Command2.Enabled = True
End Sub

'Private Sub grpid_KeyDown(K As Integer, S As Integer)
'If CHECKTB("pgr.dbf") = 1 Then
'    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'    Exit Sub
'End If
'Dim rspgr As New ADODB.Recordset
'If K = 13 Or K = 40 Or K = 38 Then
'    If rspgr.State = 1 Then rspgr.Close
'    If Len(Trim(Text2)) <> 0 Then
'        rspgr.Open "select * from pgr where grp_idy like '" & Sincrement(Trim(Text2)) & "'", CON, adOpenStatic
'        If rspgr.RecordCount > 0 Then
'            grpid = Sincrement(grpid)
'            GRPSEG = rspgr("grp_cLS")
'            GRPCAT = rspgr("GRP_CAT")
'            If K = 38 Then
'                Text6.SetFocus
'            Else
'                Text3.SetFocus
'            End If
'        Else
'            prdgrplist
'        End If
'    Else
'        prdgrplist
'   End If
'End If
'End Sub

Private Sub Text3_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsact As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If RS.State = 1 Then RS.Close
    If Len(Trim(Text3)) <> 0 Then
        rsact.Open "select * from ACT where ACT_idy like '" & Sincrement(Trim(Text3)) & "' AND (act_typ like 'PUR' OR act_typ like 'BOT')", CON, adOpenStatic
        If rsact.RecordCount > 0 Then
            Text3 = Sincrement(GRP_IDY)
            Label12 = rsact("ACT_nme")
            If K = 38 Then
                Text2.SetFocus
            Else
                Text13.SetFocus
            End If
            Command2.SetFocus
        Else
            suplist
        End If
    Else
        suplist
    End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "c" Then
    CANCELRECORD
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    
Else
    Unload Me
End If
End Sub

Public Sub CANCELRECORD()
clrctr Me
Frame2.Visible = False
Frame3.Visible = False
Frame1.Enabled = True
opt_act.Enabled = True
opt_prd.Enabled = True
opt_pgr.Enabled = True
txtFileName.SetFocus
End Sub
Private Sub arealist()
If CHECKTB("loc.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsloc As New ADODB.Recordset
If rsloc.State = 1 Then rsloc.Close
Dim ls As ListItem
rsloc.Open "select * from LOC where LOC_NME like '" & AREAID & "%'", CON, adOpenKeyset, adLockOptimistic
If rsloc.RecordCount > 0 Then
    lv.ListItems.Clear
    For A = 1 To rsloc.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(rsloc("LOC_nme")))
        ls.ListSubItems.Add , , Trim(rsloc("LOC_IDY"))
        rsloc.MoveNext
    Next
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    Label20 = "Select Location"
    lv.ColumnHeaders(1).Text = "Location Name"
    lv.ColumnHeaders(2).Text = "Location ID"
    lv.Visible = True
    lv.SetFocus
Else
    MsgBox "Location Not Found, Define Location !", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    AREAID.SetFocus
End If
End Sub
Private Sub LV_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If lv.SortOrder = lvwAscending Then
lv.SortOrder = lvwDescending
ElseIf lv.SortOrder = lvwDescending Then
lv.SortOrder = lvwAscending
End If
lv.SortKey = CH.Index - 1
End Sub

Private Sub LV_DblClick()
lv_KeyPress 13
End Sub

Private Sub lv_KeyPress(K As Integer)
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If K = 13 Then
    If lv.ListItems.Count > 0 Then
        If Label20 = "Select Location" Then
            If lv.ListItems.Count > 0 Then
                Dim lst As ListSubItem
                Set lst = lv.SelectedItem.ListSubItems.Item(1)
                If RS.State = 1 Then RS.Close
                RS.Open "select * from LOC where LOC_IDY like '" & Trim(lst.Text) & "'", CON, adOpenStatic
                AREAID = RS("LOC_IDY")
                Label29 = RS("LOC_NME")
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                CATID.SetFocus
            End If
        ElseIf Label20 = "Select Market" Then
            If lv.ListItems.Count > 0 Then
                Dim lst2 As ListSubItem
                Set lst2 = lv.SelectedItem.ListSubItems.Item(1)
                If rsmar.State = 1 Then rsmar.Close
                rsmar.Open "select * from mar where mar_idy like '" & Trim(lst2.Text) & "'", CON, adOpenStatic
                If rsmar.RecordCount > 0 Then
                    marId = rsmar("mar_idy")
                    Label30 = rsmar("mar_nme")
                    If MOD1 = False Then
                        If RSSMN.State = 1 Then RSSMN.Close
                        RSSMN.Open "SELECT * FROM SMN WHERE SMN_IDY LIKE '" & Trim(rsmar("SMN_IDY")) & "'", CON, adOpenKeyset, adLockOptimistic
                        If RSSMN.RecordCount > 0 Then Label31 = Trim(RSSMN("smn_nme"))
                        Text32 = Trim(rsmar!SMN_IDY & "")
                        rlst = Trim(rsmar!ROU_IDY & "")
                        For K = 0 To ROU_LST.ListCount - 1
                            If Trim(ROU_LST.list(K)) = Trim(rlst) Then
                                ROU_LST.ListIndex = K
                                Exit For
                            End If
                        Next
                   End If
                End If
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                Command2.SetFocus
            End If
        ElseIf Label20 = "Select Category" Then
            If lv.ListItems.Count > 0 Then
                Dim lst3 As ListSubItem
                Set lst3 = lv.SelectedItem.ListSubItems.Item(1)
                If RSCAT.State = 1 Then RS.Close
                RSCAT.Open "select * from CAT where CAT_idy like '" & Trim(lst3.Text) & "'", CON, adOpenStatic
                CATID = RSCAT("CAT_idy")
                Label19 = RSCAT("CAT_nme")
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                marId.SetFocus
            End If
        ElseIf Label20 = "Select Product Category" Then
            If lv.ListItems.Count > 0 Then
                grpid = lv.SelectedItem.ListSubItems(1)
                GRPSEG = lv.SelectedItem.ListSubItems(2)
                GRPCAT = lv.SelectedItem
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                Text3.SetFocus
            End If
        ElseIf Label20 = "Select Supplier" Then
            If lv.ListItems.Count > 0 Then
                Text3 = lv.SelectedItem.ListSubItems(1)
                Label12 = lv.SelectedItem
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                Command2.SetFocus
            End If
        ElseIf Label20 = "Select Group Account" Then
            If lv.ListItems.Count > 0 Then
                GRP_IDY = lv.SelectedItem.ListSubItems(1)
                Label4 = lv.SelectedItem
                SSTab1.TabEnabled(1) = False
                lv.Visible = False
                SSTab1.Tab = 0
                AREAID.SetFocus
            End If
        End If
    End If
End If
End Sub

Private Sub list4_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub CATID_KeyDown(K As Integer, S As Integer)
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim RSCAT As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If RSCAT.State = 1 Then RSCAT.Close
        RSCAT.Open " select * from CAT where CAT_idy like '" & Sincrement(CATID) & "'", CON, adOpenStatic
        If RSCAT.RecordCount > 0 Then
            CATID = Sincrement(CATID)
            Label19 = RSCAT("CAT_nme")
            If K = 38 Then
                AREAID.SetFocus
            Else
                marId.SetFocus
            End If
        Else
            chanlist
        End If
End If
End Sub

Private Sub catid_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Or KeyAscii = 37 Then KeyAscii = 0
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub chanlist()
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim RSCAT As New ADODB.Recordset
If RSCAT.State = 1 Then RS.Close
RSCAT.Open "select * from CAT where CAT_nme like '" & CATID & "%'", CON, adOpenKeyset, adLockOptimistic
If RSCAT.RecordCount > 0 Then
    lv.ListItems.Clear
    For A = 1 To RSCAT.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(RSCAT("CAT_nme")))
        ls.ListSubItems.Add , , Trim(RSCAT("CAT_idy"))
        RSCAT.MoveNext
    Next
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    Label20 = "Select Category"
    lv.ColumnHeaders(1).Text = "Category Name"
    lv.ColumnHeaders(2).Text = "Category ID"
    lv.Visible = True
    lv.SetFocus
Else
    MsgBox "Category Not Found, Define Category !", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    CATID.SetFocus
End If
End Sub

Private Sub MARID_KeyPress(K As Integer)
 If K > 96 And K < 123 Then K = K - 32
 If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub MARID_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("smn.dbf,mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsmar As New ADODB.Recordset
Dim RSSMN As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If rsmar.State = 1 Then rsmar.Close
        rsmar.Open " select * from mar where mar_idy like '" & Sincrement(marId) & "'", CON, adOpenStatic
        If rsmar.RecordCount > 0 Then
            marId = Sincrement(marId)
            Label30 = rsmar("mar_nme")
            If MOD1 = False Then
                Text32 = Trim(rsmar!SMN_IDY)
                If RSSMN.State = 1 Then RSSMN.Close
                RSSMN.Open "SELECT SMN_NME FROM SMN WHERE SMN_IDY LIKE '" & Trim(rsmar("SMN_IDY")) & "'", CON, adOpenKeyset, adLockOptimistic
                If RSSMN.RecordCount > 0 Then Label31 = RSSMN("smn_nme")
                'ROU_LST = Trim(rsmar!ROU_IDY)
                rlst = Trim(rsmar!ROU_IDY & "")
                For K = 0 To ROU_LST.ListCount - 1
                    If Trim(ROU_LST.list(K)) = Trim(rlst) Then
                        ROU_LST.ListIndex = K
                        Exit For
                    End If
                Next
            Else
                
            End If
            If K = 38 Then
                CATID.SetFocus
            Else
                Text32.SetFocus
            End If
        Else
            marlist
        End If
End If
End Sub

Private Sub marlist()
If CHECKTB("mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsmar As New ADODB.Recordset
If rsmar.State = 1 Then rsmar.Close
rsmar.Open "select * from mar where mar_nme like '" & Trim(marId) & "%'", CON, adOpenKeyset, adLockOptimistic
If rsmar.RecordCount > 0 Then
    lv.ListItems.Clear
    For A = 1 To rsmar.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(rsmar("mar_nme")))
        ls.ListSubItems.Add , , Trim(rsmar("mar_idy"))
        rsmar.MoveNext
    Next
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    Label20 = "Select Market"
    lv.ColumnHeaders(1).Text = "Market Name"
    lv.ColumnHeaders(2).Text = "Market ID"
    lv.Visible = True
    lv.SetFocus
Else
    MsgBox "Market Not Found, Define Market !", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    marId.SetFocus
End If
End Sub

Private Sub prdgrplist()
If CHECKTB("pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rspgr As New ADODB.Recordset
If rspgr.State = 1 Then rspgr.Close
rspgr.Open "select * from pgr where grp_IDY like '" & Trim(grpid) & "%'", CON, adOpenKeyset
If rspgr.RecordCount > 0 Then
    lv.ListItems.Clear
    lv.ColumnHeaders.Clear
    lv.ColumnHeaders.Add , , "Product Category"
    lv.ColumnHeaders.Add , , "ID"
    lv.ColumnHeaders.Add , , "Product Segment"
    For A = 1 To rspgr.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(rspgr("grp_cat")))
        ls.ListSubItems.Add , , Trim(rspgr("grp_idy"))
        ls.ListSubItems.Add , , Trim(rspgr("GRP_CLS") & "")
        rspgr.MoveNext
    Next
    SSTab1.Tab = 1
    Label31 = "Select Product Category"
    lv.Visible = True
    lv.SetFocus
    Label20 = "Select Product Category"
Else
    MsgBox "Segment Not Found, Define Segment !", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    grpid.Enabled = True
    grpid.SetFocus
End If
End Sub

Private Sub suplist()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rsact As New ADODB.Recordset
If rsact.State = 1 Then rsact.Close
'If Check3.Value = 1 Then
'rs.Open "select * from ACT where ACT_nme like '" & Trim(Text3) & "%' and (act_typ like 'PUR' or act_typ like 'BOT')", CON, adOpenKeyset
'ElseIf Check4.Value = 1 Then
rsact.Open "select * from ACT where act_ad4 like '" & Trim(Text3) & "%' and (act_typ like 'PUR' OR act_typ like 'BOT')", CON, adOpenKeyset
'End If
lv.ListItems.Clear
lv.ColumnHeaders.Clear
lv.ColumnHeaders.Add , , "Supplier Name"
lv.ColumnHeaders.Add , , "ID"
lv.ColumnHeaders.Add , , "Place"
Label20 = "Select Supplier"
If rsact.RecordCount > 0 Then
    For A = 1 To rsact.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(rsact("ACT_nme")))
        ls.ListSubItems.Add , , Trim(rsact("ACT_idy"))
        ls.ListSubItems.Add , , Trim(rsact("act_ad4") & "")
        rsact.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Supplier"
    lv.Visible = True
    lv.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Supplier Not Found, Define Supplier !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        Text3.Enabled = True
        Text3.SetFocus
    End If
End If
End Sub

Private Sub txtFileName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command1.SetFocus
End If
End Sub
Private Sub SHOWBUT(b As Boolean)
Label5.Visible = b
Label6.Visible = b
Label7.Visible = b
Label8.Visible = b
End Sub
'Private Sub prdgrplist()
'If CHECKTB("pgr.dbf") = 1 Then
'    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'    cancel
'    Exit Sub
'End If
'Dim rs As New ADODB.Recordset
'If Check1.Value = 1 Then
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from pgr where grp_cat like '" & text1 & "%'", CON, adOpenKeyset, adLockOptimistic
'Else
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from pgr where grp_cls like '" & text1 & "%'", CON, adOpenKeyset, adLockOptimistic
'End If
'list1.ListItems.Clear
'If rs.RecordCount > 0 Then
'    For A = 1 To rs.RecordCount
'        Set ls = list1.ListItems.Add(, , Trim(rs("GRP_CAT")))
'        ls.ListSubItems.Add , , Trim(rs("GRP_CLS"))
'        ls.ListSubItems.Add , , Trim(rs("GRP_IDY"))
'        rs.MoveNext
'    Next
'    SSTab1.Tab = 1
'    list1.Visible = True
'    list1.SetFocus
'Else
'    MsgBox "Category Not Found, Define Category !", vbExclamation, "Effmcg"
'    SSTab1.Tab = 0
'    text1.Enabled = True
'    text1.SetFocus
'End If
'End Sub
Private Function ADS(RS As ADODB.Recordset, FLD As String, DB As String)
If RS(FLD).DefinedSize < Len(Trim(DB)) Then
    ADS = Mid(DB, 1, RS(FLD).DefinedSize)
Else
    ADS = Trim(DB)
End If
End Function

