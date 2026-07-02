VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab2 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Products"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&View"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label17"
      Tab(1).Control(1)=   "Lst1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame7 
         Height          =   4335
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   8175
         Begin VB.CommandButton Command9 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   6960
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   3840
            Width           =   975
         End
         Begin VB.CommandButton Command8 
            Caption         =   "&OK"
            Height          =   375
            Left            =   5880
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   3840
            Width           =   975
         End
         Begin VB.CommandButton Command10 
            Caption         =   "&Select All"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   3840
            Width           =   1575
         End
         Begin MSComctlLib.ListView LVMOD 
            Height          =   3405
            Left            =   240
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   360
            Width           =   7665
            _ExtentX        =   13520
            _ExtentY        =   6006
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
      End
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   8175
         Begin VB.TextBox sal_smn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            ForeColor       =   &H80000005&
            Height          =   330
            Left            =   2400
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox sal_idy 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox rou_idy 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5160
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblsalesman 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Salesman"
            Height          =   195
            Left            =   360
            TabIndex        =   6
            Top             =   315
            Width           =   690
         End
         Begin VB.Label lblroute1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Route"
            Height          =   195
            Left            =   4575
            TabIndex        =   5
            Top             =   240
            Width           =   435
         End
      End
      Begin MSComctlLib.ListView Lst1 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   12
         Top             =   960
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7435
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "Order ID"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order Date"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Salesman"
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
         Height          =   375
         Left            =   -74880
         TabIndex        =   13
         Top             =   600
         Width           =   8535
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
