VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_SMN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distributor - Dealer Salesmen"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "M_FRM_SMN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
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
            Picture         =   "M_FRM_SMN.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SMN.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SMN.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SMN.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SMN.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SMN.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SMN.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SMN.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SMN.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SMN.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SMN.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SMN.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   435
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   9128
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
      TabCaption(0)   =   "&Salesmen Details"
      TabPicture(0)   =   "M_FRM_SMN.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label14"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text10"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text9"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text8"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text7"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text6"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text5"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Frame2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtmcldep"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_SMN.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "list1"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtmcldep 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6240
         MaxLength       =   9
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   915
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Invoice Credit Check"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   5040
         TabIndex        =   27
         Top             =   1305
         Width           =   2895
         Begin VB.TextBox txtmcldays 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtmclAmt 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   1200
            MaxLength       =   9
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   975
            Width           =   1335
         End
         Begin VB.ListBox lstMcl 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            ItemData        =   "M_FRM_SMN.frx":3386
            Left            =   1200
            List            =   "M_FRM_SMN.frx":3390
            TabIndex        =   30
            Top             =   255
            Width           =   735
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Days"
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
            Left            =   660
            TabIndex        =   35
            Top             =   645
            Width           =   360
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
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
            Left            =   480
            TabIndex        =   29
            Top             =   1050
            Width           =   540
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Required"
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
            Left            =   375
            TabIndex        =   28
            Top             =   285
            Width           =   645
         End
      End
      Begin VB.TextBox Text3 
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
         Height          =   315
         Left            =   1905
         MaxLength       =   25
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1335
         Width           =   2805
      End
      Begin VB.TextBox Text2 
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
         Left            =   1905
         MaxLength       =   40
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   915
         Width           =   2805
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   -74640
         TabIndex        =   21
         Top             =   4560
         Width           =   3015
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   120
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Place"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   480
         Width           =   1035
      End
      Begin VB.TextBox Text4 
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
         Left            =   1905
         MaxLength       =   25
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1695
         Width           =   2805
      End
      Begin VB.TextBox Text5 
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
         Left            =   1905
         MaxLength       =   25
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2805
      End
      Begin VB.TextBox Text6 
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
         Left            =   1905
         MaxLength       =   25
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2415
         Width           =   2805
      End
      Begin VB.TextBox Text7 
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
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2835
         Width           =   1035
      End
      Begin VB.TextBox Text8 
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
         Left            =   1905
         MaxLength       =   25
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3270
         Width           =   2805
      End
      Begin VB.TextBox Text9 
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
         Left            =   1905
         MaxLength       =   25
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3705
         Width           =   2805
      End
      Begin VB.TextBox Text10 
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
         Left            =   1905
         MaxLength       =   25
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   4125
         Width           =   2835
      End
      Begin VB.TextBox Text11 
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
         Left            =   1905
         MaxLength       =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4560
         Width           =   2835
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3495
         Left            =   -74640
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   14208
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
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Deposit"
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
         Left            =   5520
         TabIndex        =   33
         Top             =   990
         Width           =   540
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
         Left            =   -74640
         TabIndex        =   20
         Top             =   600
         Width           =   8055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salesman ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   660
         TabIndex        =   18
         Top             =   548
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1140
         TabIndex        =   17
         Top             =   983
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   990
         TabIndex        =   16
         Top             =   1395
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   795
         TabIndex        =   15
         Top             =   3338
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1305
         TabIndex        =   14
         Top             =   3773
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile/Pager #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   450
         TabIndex        =   13
         Top             =   4193
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1185
         TabIndex        =   12
         Top             =   4628
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City/Place"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   840
         TabIndex        =   11
         Top             =   2483
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zip/Postal Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   405
         TabIndex        =   10
         Top             =   2903
         Width           =   1155
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   8850
      _ExtentX        =   15610
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
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Salesmen ID "
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_SMN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD1 As Boolean, idnew As Boolean
Private Sub CREATE()
MOD1 = False
Text1 = ""
save
End Sub

Public Sub delete()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

   If MsgBox("Delete Salesman ?", vbYesNo + vbCritical, "DAS Version 4.0") = vbYes Then
        If check("smn", "smn_idy", Text1) = False Then
            MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
            SSTab1.Tab = 0
            clrctr M_FRM_SMN
            Text1.Enabled = True
            Text1.SetFocus
            Exit Sub
        Else
            If deletecheck = False Then
                con.Execute "delete from smn where smn_idy like '" & Sincrement(Text1) & "'"
                SSTab1.Tab = 0
                Text1.Enabled = True
                Text1.SetFocus
                clrctr M_FRM_SMN
            Else
                MsgBox "Cannot Delete, Transaction Exists !", vbExclamation, "DAS Version 4.0"
                Text2.SetFocus
                Exit Sub
            End If
        End If
        cancel
    Else
        cancel
    End If
End Sub
Public Function deletecheck() As Boolean
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Function
End If

If check("act", "smn_idy", Trim(Text1)) = False Then
    If check("inl", "smn_idy", Trim(Text1)) = False Then
        deletecheck = False
    Else
        deletecheck = True
    End If
Else
    deletecheck = True
End If
End Function

Private Sub save()
Dim rs As New ADODB.Recordset
If MsgBox("Confirm Record Save ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If Len(Trim(Text2)) = 0 Then
    MsgBox "Blank Name, Enter Name !", vbExclamation, "DAS Version 4.0"
    Text2.SetFocus
    Exit Sub
End If
If rs.State = 1 Then rs.Close
If MOD1 = True Then
rs.Open "select * from smn where smn_idy like '" & Sincrement(Text1) & "'", con, adOpenKeyset, adLockPessimistic
Else
rs.Open "select * from smn", con, adOpenKeyset, adLockPessimistic
rs.AddNew
rs("smn_idy") = increment("idy_smn")
MsgBox "New Salesmen ID is " & rs("smn_idy"), vbExclamation, "DAS Version 4.0"
End If
rs("smn_nme") = Trim(Text2)
rs("smn_ad1") = Trim(Text3)
rs("smn_ad2") = Trim(Text4)
rs("smn_ad3") = Trim(Text5)
rs("smn_ad4") = Trim(Text6)
rs("smn_tfx") = Trim(Text8)
rs("smn_fax") = Trim(Text9) & ""
rs("smn_mbl") = Trim(Text10)
rs("smn_eml") = Trim(Text11)
rs("smn_zip") = Val(Text7)
rs("smn_dep") = Val(txtmcldep)
rs("chk_flg") = Trim(lstMcl.Text)
rs("smn_amt") = Val(txtmclAmt.Text)
rs("smn_dys") = Val(txtmcldays)
rs.Update
If rs.State = 1 Then rs.Close
clrctr M_FRM_SMN
idnew = False
SSTab1.Tab = 0
Text1.Enabled = True
Text1.SetFocus
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub cancel()
clrctr M_FRM_SMN
idnew = False
SSTab1.Tab = 0
Text1.Enabled = True
Text1.SetFocus
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub newrec()
MOD1 = False
clrctr M_FRM_SMN
Text1.Enabled = False
Text2.SetFocus
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub smnlist()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

Dim ls As ListItem
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Check1.Value = 1 Then
    rs.Open "select * from smn where smn_nme like '" & Text1 & "%'", con, adOpenStatic
ElseIf Check2.Value = 1 Then
    rs.Open "select * from smn where smn_ad4 like '" & Text1 & "%'", con, adOpenStatic
End If
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("smn_nme")))
        ls.ListSubItems.Add , , Trim(rs("smn_idy") & "")
        ls.ListSubItems.Add , , Trim(rs("smn_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.ListItems(1).Selected = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text1.Enabled = True
        Text1.SetFocus
    End If
End If
End Sub

Private Sub exitfrm()
Unload Me
Set Form = Nothing
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
    smnlist
Else
    Check2.Value = 1
    Check2_Click
End If
SSTab1.TabEnabled(1) = False
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
    smnlist
Else
    Check1.Value = 1
    Check1_Click
End If
SSTab1.TabEnabled(1) = False
End Sub

Private Sub Check2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Place To Get List of Salesmen Sorted on Place"
End Sub

Private Sub Check1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Name To Get List of Salesmen Sorted on Name"
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
If Text1.Enabled = True Then
    Text1.SetFocus
Else
    If Text2.Enabled = True Then Text2.SetFocus
End If
txtmclAmt = "0.00"
txtmcldep = "0.00"
lstMcl.ListIndex = 1
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_SMN.ActiveControl Is TextBox Then M_FRM_SMN.ActiveControl.Text = Trim(FUNKEY(K))
If list1.Visible = False Then
    If K = 78 And s = 2 Then
        newrec
    ElseIf K = 83 And s = 2 Then
        save
    ElseIf K = 81 And s = 2 Then
        cancel
    ElseIf K = 88 And s = 2 Then
        exitfrm
    ElseIf K = 67 And s = 2 Then
        If Toolbar1.Buttons(12).Enabled = True Then CREATE
    ElseIf K = 68 And s = 2 Then
        delete
    ElseIf K = 69 And s = 2 Then
        If M_FRM_SMN.ActiveControl.name = "Text1" Then smnlist
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And s = 2 Then
        OpenReport CURDIR & "\lists.RP1", "Salesmen list", 0, "", 6
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
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        If CHECKTB("smn.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            cancel
            Exit Sub
        End If

        If check("smn", "smn_idy", Trim(Text1)) = False Then
            Text1.SetFocus
        Else
            Text1.Enabled = False
            Text2.SetFocus
        End If
        Toolbar1.Enabled = True
    Else
        If M_FRM_SMN.ActiveControl.name = "Text1" Then
            Unload Me
            Set Form = Nothing
        ElseIf M_FRM_SMN.ActiveControl.name = "SSTab1" Then
        Else
            save
        End If
    End If
End If
End Sub


Private Sub Form_Load()
checkdongle
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
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

Private Sub list1_KeyDown(KeyAscii As Integer, s As Integer)
Dim rs As New ADODB.Recordset
If KeyAscii = 13 Then
    If list1.ListItems.count > 0 Then
        If rs.State = 1 Then rs.Close
        Dim lst As ListSubItem
        Set lst = list1.SelectedItem.ListSubItems.Item(1)
        rs.Open "select * from smn where smn_idy like '" & Trim(lst.Text) & "'", con, adOpenStatic
        Text1 = Trim(rs("smn_idy"))
        Text2 = Trim(rs("smn_nme"))
        Text3 = Trim(rs("smn_ad1") & "")
        Text4 = Trim(rs("smn_ad2") & "")
        Text5 = Trim(rs("smn_ad3") & "")
        Text6 = Trim(rs("smn_ad4") & "")
        Text8 = Trim(rs("smn_tfx") & "")
        Text9 = Trim(rs("smn_fax") & "")
        Text10 = Trim(rs("smn_mbl") & "")
        Text11 = Trim(rs("smn_eml") & "")
        Text7 = Trim(rs("smn_zip") & "")
        If Trim(rs("chk_flg")) = "Y" Then
            lstMcl.Selected(0) = True
        Else
            lstMcl.Selected(1) = True
        End If
        txtmclAmt = convert(Val(Trim(rs("smn_amt")) & ""))
        txtmcldep = convert(Val(Trim(rs("smn_dep")) & ""))
        txtmcldays = Val(Trim(rs("smn_dys")) & "")
        list1.Visible = False
        SSTab1.Tab = 0
        Text2.SetFocus
        Text1.Enabled = False
        MOD1 = True
        Toolbar1.Enabled = True
        Toolbar1.Buttons(12).Enabled = True
    End If
End If
End Sub

Private Sub lstMcl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Trim(lstMcl.Text) = "Y" Then
        txtmcldays.Enabled = True
        txtmclAmt.Enabled = True
        'txtmcldep.Enabled = True
        txtmcldays.SetFocus
        SendKeys "{HOME}+{END}"
    Else
        txtmcldays.Enabled = False
        txtmclAmt.Enabled = False
        'txtmcldep.Enabled = False
        save
    End If
ElseIf KeyCode = 38 And Shift = 1 Then
    Text11.SetFocus
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Toolbar1.Enabled = True
list1.Visible = False
If SSTab1.Tab = 0 Then
    If Text1.Enabled = True Then
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Salesmen ID or Press <Enter> To Get The List of Salesmen ID's"
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 27 And Text1 = "" Then Unload Me
Set Form = Nothing
End Sub

Private Sub Text10_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Mobile or Pager Number"
End Sub

Private Sub Text11_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Email ID"
End Sub

Private Sub Text13_Change()

End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Salesmen Name"
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Address"
End Sub

Private Sub text3_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text4_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text5_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text6_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter City/Place"
End Sub

Private Sub Text6_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Text7_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Postal or Zip Code"
End Sub

Private Sub TEXT1_KeyDown(K As Integer, Shift As Integer)
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from smn where smn_idy like '" & Sincrement(Text1) & "'", con, adOpenStatic
    If rs.RecordCount > 0 Then
        Text1 = Sincrement(Text1)
        Text2 = Trim(rs("smn_nme"))
        Text3 = Trim(rs("smn_ad1") & "")
        Text4 = Trim(rs("smn_ad2") & "")
        Text5 = Trim(rs("smn_ad3") & "")
        Text6 = Trim(rs("smn_ad4") & "")
        Text7 = Trim(rs("smn_zip") & "")
        Text9 = Trim(rs("smn_fax") & "")
        Text10 = Trim(rs("smn_mbl") & "")
        Text11 = Trim(rs("smn_eml") & "")
        Text8 = Trim(rs("smn_tfx") & "")
        If Trim(rs("chk_flg")) = "Y" Then
            lstMcl.Selected(0) = True
        Else
            lstMcl.Selected(1) = True
        End If
        txtmclAmt = convert(Val(Trim(rs("smn_amt")) & ""))
        txtmcldep = convert(Val(Trim(rs("smn_dep")) & ""))
        txtmcldays = Trim(rs("smn_dys")) & ""
        MOD1 = True
        Text1.Enabled = False
        Text2.SetFocus
        Toolbar1.Buttons(12).Enabled = True
    Else
        smnlist
    End If
End If
End Sub

Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Text2)) > 0 Then
        Text3.SetFocus
    Else
        MsgBox "Blank Name, Enter Name !", vbExclamation, "DAS Version 4.0"
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text4.SetFocus
ElseIf KeyCode = 38 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text5.SetFocus
ElseIf KeyCode = 38 Then
    Text3.SetFocus
End If
End Sub

Private Sub TEXT5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text6.SetFocus
ElseIf KeyCode = 38 Then
    Text4.SetFocus
End If
End Sub

Private Sub TEXT6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text7.SetFocus
ElseIf KeyCode = 38 Then
    Text5.SetFocus
End If
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text8.SetFocus
ElseIf KeyCode = 38 Then
    Text6.SetFocus
End If
End Sub

Private Sub Text8_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Telephone Number"
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text9.SetFocus
ElseIf KeyCode = 38 Then
    Text7.SetFocus
End If
End Sub

Private Sub Text8_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Text9_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Fax Number"
End Sub

Private Sub Text9_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text10.SetFocus
ElseIf KeyCode = 38 Then
    Text8.SetFocus
End If
End Sub

Private Sub Text10_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text11.SetFocus
ElseIf KeyCode = 38 Then
    Text9.SetFocus
End If
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtmcldep.SetFocus
    SendKeys "{home}+{end}"
ElseIf KeyCode = 38 Then
    Text10.SetFocus
End If
End Sub

Private Sub Text7_keypress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    newrec
ElseIf Button.KEY = "s" Then
    save
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "e" Then
    exitfrm
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If M_FRM_SMN.ActiveControl.name = "Text1" Then smnlist
ElseIf Button.KEY = "d" Then
    delete
ElseIf Button.KEY = "pre" Then
    SYS_PSW.Tag = "SMN_PRE"
    SYS_PSW.Show 1
ElseIf Button.KEY = "p" Then
    OpenReport CURDIR & "\lists.RP1", "Salesmen list", 0, "", 6
    cancel
ElseIf Button.KEY = "pr" Then
ElseIf Button.KEY = "crt" Then
    CREATE
End If

End Sub

Private Sub txtmclAmt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtmclAmt = convert(txtmclAmt)
    save
ElseIf KeyCode = 38 Then
    lstMcl.SetFocus
End If
End Sub

Private Sub txtmclAmt_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtmcldays_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Val(txtmcldays) <= 365 Then
        txtmclAmt.SetFocus
        SendKeys "{home}+{end}"
    Else
        MsgBox "Valid Days <= 365 !", vbExclamation, "DAS Version 4.0"
        txtmcldays.SetFocus
        SendKeys "{home}+{end}"
    End If
ElseIf KeyCode = 38 Then
    txtmclAmt.SetFocus
End If
End Sub

Private Sub txtmcldays_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtmcldep_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    lstMcl.SetFocus
ElseIf KeyCode = 38 Then
    txtmcldays.SetFocus
End If
End Sub

Private Sub txtmcldep_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
Else
    KeyAscii = 0
End If
End Sub
