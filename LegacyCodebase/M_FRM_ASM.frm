VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_ASM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assets"
   ClientHeight    =   5625
   ClientLeft      =   -330
   ClientTop       =   1440
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "M_FRM_ASM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
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
            Picture         =   "M_FRM_ASM.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ASM.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ASM.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ASM.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ASM.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ASM.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ASM.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ASM.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ASM.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ASM.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ASM.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ASM.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5140
      Left            =   0
      TabIndex        =   1
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
      TabCaption(0)   =   "&Asset Details"
      TabPicture(0)   =   "M_FRM_ASM.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "M_AST_NME"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "M_MFG_NME"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "M_MNT_NME"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label12"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "M_AST_TYP"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "M_AST_VAL"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "M_AST_DEP"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "M_AST_LFP"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "M_MNT_IDY"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "M_MFG_IDY"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "M_CMP_IDY"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "M_AST_RMK"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "indidy"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "M_AST_CAP"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_ASM.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "list1"
      Tab(1).ControlCount=   2
      Begin VB.TextBox M_AST_CAP 
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
         Height          =   330
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ListBox indidy 
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
         ItemData        =   "M_FRM_ASM.frx":3386
         Left            =   2460
         List            =   "M_FRM_ASM.frx":3390
         TabIndex        =   27
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox M_AST_RMK 
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
         Height          =   690
         Left            =   2460
         MultiLine       =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox M_CMP_IDY 
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
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox M_MFG_IDY 
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
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox M_MNT_IDY 
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
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox M_AST_LFP 
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
         Height          =   330
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox M_AST_DEP 
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
         Height          =   330
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3300
         Width           =   1095
      End
      Begin VB.TextBox M_AST_VAL 
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
         Height          =   330
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2940
         Width           =   1095
      End
      Begin VB.TextBox M_AST_TYP 
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
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
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
         Left            =   2460
         MaxLength       =   35
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   3495
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
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1065
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1545
         TabIndex        =   30
         Top             =   2595
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Individual ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1275
         TabIndex        =   28
         Top             =   3675
         Width           =   885
      End
      Begin VB.Label M_MNT_NME 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3600
         TabIndex        =   26
         Top             =   2160
         Width           =   2340
      End
      Begin VB.Label M_MFG_NME 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3600
         TabIndex        =   24
         Top             =   1800
         Width           =   2340
      End
      Begin VB.Label M_AST_NME 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3600
         TabIndex        =   23
         Top             =   1080
         Width           =   2340
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1575
         TabIndex        =   22
         Top             =   4320
         Width           =   630
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Identification"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   555
         TabIndex        =   21
         Top             =   1515
         Width           =   1605
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturer ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1005
         TabIndex        =   19
         Top             =   1875
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maintained By ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   945
         TabIndex        =   17
         Top             =   2235
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Life Period (Months)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   750
         TabIndex        =   15
         Top             =   4035
         Width           =   1410
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Depreciation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1260
         TabIndex        =   13
         Top             =   3368
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1755
         TabIndex        =   11
         Top             =   3008
         Width           =   405
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asset Type ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1155
         TabIndex        =   9
         Top             =   1155
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
         Left            =   -74520
         TabIndex        =   6
         Top             =   720
         Width           =   7900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asset ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1560
         TabIndex        =   5
         Top             =   435
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asset Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1305
         TabIndex        =   4
         Top             =   795
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
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
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AST_CID"
                  Text            =   "Assets - Company"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AST_MFG"
                  Text            =   "Assets - Manufacturer"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AST_REC"
                  Text            =   "Assets Movement Record"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AST_POS"
                  Text            =   "Assets Position"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AST_TRK"
                  Text            =   "Asset Tracking"
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
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Assets ID"
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
Attribute VB_Name = "M_FRM_ASM"
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
If CHECKTB("asm.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If MsgBox("Are you sure you want to delete Asset ?", vbYesNo + vbCritical, "DAS Version 4.0") = vbYes Then
    If check("ASM", "ASM_idy", Text1) = False Then
        MsgBox "Asset Not Found, Define Asset !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        clrctr M_FRM_ASM
        Text1.Enabled = True
        Text1.SetFocus
        Exit Sub
    Else
        If deletecheck = False Then
            con.Execute "delete from ASM where ASM_idy like '" & Sincrement(Text1) & "'"
            SSTab1.Tab = 0
            Text1.Enabled = True
            Text1.SetFocus
            clrctr M_FRM_ASM
        Else
            MsgBox "Cannot Delete, Transaction Exists", vbExclamation, "DAS Version 4.0"
            Text2.SetFocus
            Exit Sub
        End If
    End If
    cancel
Else
    cancel
End If
End Sub
Private Function deletecheck() As Boolean
If CHECKTB("ast.dbf,prt.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Function
End If
If check("aST", "ASM_idy", Trim(Text1)) = False Then
    deletecheck = False
Else
    deletecheck = True
End If
End Function
Private Sub save()
On Error GoTo lab
Dim rs As New ADODB.Recordset
If MsgBox("Confirm Record Save ?", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
If CHECKTB("asm.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If Len(Trim(Text2)) = 0 Then
    MsgBox "Asset Name Should Not Be Blank, Give Name !", vbExclamation, "DAS Version 4.0"
    Text2.Enabled = True
    Text2.SetFocus
    Exit Sub
End If
If Len(Trim(M_AST_TYP)) = 0 Then
    MsgBox "Asset ID Should not be Blank!", vbExclamation, "DAS Version 4.0"
    M_AST_TYP.Enabled = True
    M_AST_TYP.SetFocus
    Exit Sub
End If
If Len(Trim(M_MFG_IDY)) = 0 Then
    MsgBox "Manufacturer ID Should not be Blank!", vbExclamation, "DAS Version 4.0"
    M_MFG_IDY.Enabled = True
    M_MFG_IDY.SetFocus
    Exit Sub
End If
If Len(Trim(M_MNT_IDY)) = 0 Then
    MsgBox "Maintained By ID Should not be Blank!", vbExclamation, "DAS Version 4.0"
    M_MNT_IDY.Enabled = True
    M_MNT_IDY.SetFocus
    Exit Sub
End If
If rs.State = 1 Then rs.Close
If MOD1 = True Then
rs.Open "select * from ASM where ASM_idy like '" & Sincrement(Text1) & "'", con, adOpenKeyset, adLockPessimistic
Else
rs.Open "select * from ASM", con, adOpenKeyset, adLockPessimistic
rs.AddNew
rs("ASM_idy") = increment("idy_asm")
MsgBox "New Asset ID is " & rs("ASM_idy"), vbExclamation, "DAS Version 4.0"
End If
rs("ASM_nme") = Trim(Text2)
rs!ATT_IDY = Trim(M_AST_TYP)
rs!ASM_VAL = Val(M_AST_VAL)
rs!ACT_IDY = Trim(M_MNT_IDY)
rs!ASM_DPV = Val(M_AST_DEP)
rs!ASM_CID = Trim(M_CMP_IDY)
rs!ASM_CAP = Val(M_AST_CAP)
rs!ASM_LFP = Val(Trim(M_AST_LFP) & "")
rs!MFG_IDY = Trim(M_MFG_IDY)
rs!ASM_TSP = Trim(M_AST_RMK)
rs!ind_idy = indidy
rs!ASM_QTY = 0
rs.Update
cancel
Text1.Enabled = True
idnew = False
SSTab1.Tab = 0
Text1.SetFocus
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
Exit Sub
lab:
MsgBox err.Description
End Sub

Private Sub cancel()
clrctr M_FRM_ASM
Text1.Enabled = True
M_AST_NME = ""
M_MFG_NME = ""
M_MNT_NME = ""
idnew = False
SSTab1.Tab = 0
Text1.SetFocus
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
indidy.ListIndex = 1
End Sub

Private Sub newrec()
MOD1 = False
clrctr M_FRM_ASM
Text1.Enabled = False
SSTab1.Tab = 0
Text2.SetFocus
Toolbar1.Buttons(12).Enabled = False
indidy.ListIndex = 1
End Sub

Private Sub Assetlist()
If CHECKTB("asm.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from ASM where ASM_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add , , "Name"
    list1.ColumnHeaders.Add , , "ID"
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("ASM_nme")))
        ls.ListSubItems.Add , , Trim(rs("ASM_idy"))
        rs.MoveNext
    Next
    Label20.Caption = "Select Asset"
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "Asset Not Found, Define Asset !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub exitfrm()
Unload Me
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
If Text1.Enabled = True Then
    Text1.SetFocus
Else
    Text2.SetFocus
End If
indidy.ListIndex = 1
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_ASM.ActiveControl Is TextBox Then M_FRM_ASM.ActiveControl.Text = Trim(FUNKEY(K))
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
        If M_FRM_ASM.ActiveControl.name = "Text1" Then Assetlist
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And s = 2 Then
        'Openreport CURDIR & "\Assets.RP1", "Assets", 0, "", 6
    ElseIf K = 82 And s = 2 Then
        'MDIForm1.RRReport1.ReportName = App.Path & "\Channels list.rrw"
        'MDIForm1.RRReport1.Destination = 1
        'MDIForm1.RRReport1.Printer = "?"
        'MDIForm1.RRReport1.RunReport 1
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        If CHECKTB("asm.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            cancel
            Exit Sub
        End If

        If Label20.Caption = "Select Asset" Then
            list1.Visible = False
            SSTab1.Tab = 0
            SSTab1.TabEnabled(1) = False
            If check("ASM", "ASM_idy", Trim(Text1)) = False Then
                Text1.Enabled = True
                Text1.SetFocus
            Else
                Text1.Enabled = False
                Text2.SetFocus
            End If
        ElseIf Label20.Caption = "Select Asset Type" Then
            SSTab1.Tab = 0
            list1.Visible = False
            M_AST_TYP.SetFocus
        ElseIf Label20.Caption = "Select Manufacturer" Then
            SSTab1.Tab = 0
            list1.Visible = False
            M_MFG_IDY.SetFocus
        ElseIf Label20.Caption = "Select Maintenance" Then
            SSTab1.Tab = 0
            list1.Visible = False
            M_MFG_IDY.SetFocus
        End If
    Else
        If M_FRM_ASM.ActiveControl.name = "Text1" Then
            Unload Me
        ElseIf M_FRM_ASM.ActiveControl.name = "SSTab1" Then
        Else
            save
        End If
    End If
End If
End Sub

Private Sub Form_Load()
checkdongle
MOD1 = False
idnew = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub indidy_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    M_AST_LFP.SetFocus
ElseIf K = 38 And Shift = 1 Then
    M_AST_DEP.SetFocus
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
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Asset or Press <Esc> to Quit"
End Sub

Private Sub list1_KeyDown(K As Integer, s As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Then
    If list1.ListItems.count > 0 Then
        If Label20.Caption = "Select Asset" Then
            Dim lst As ListSubItem
            Text1 = list1.SelectedItem.ListSubItems.Item(1)
            list1.Visible = False
            SSTab1.Tab = 0
            TEXT1_KeyDown 13, 1
            MOD1 = True
        ElseIf Label20.Caption = "Select Asset Type" Then
            M_AST_TYP = list1.SelectedItem.ListSubItems.Item(1)
            M_AST_NME = list1.SelectedItem
            list1.Visible = False
            SSTab1.Tab = 0
            M_CMP_IDY.SetFocus
        ElseIf Label20.Caption = "Select Manufacturer" Then
            M_MFG_IDY = list1.SelectedItem.ListSubItems(1)
            M_MFG_NME = list1.SelectedItem
            list1.Visible = False
            SSTab1.Tab = 0
            M_MNT_IDY.SetFocus
        ElseIf Label20.Caption = "Select Maintenance" Then
            M_MNT_IDY = list1.SelectedItem.ListSubItems(1)
            M_MNT_NME = list1.SelectedItem
            list1.Visible = False
            SSTab1.Tab = 0
            M_AST_CAP.SetFocus
        ElseIf Label20 = "Select Asset" Then
            list1.Visible = False
            SSTab1.Tab = 0
            Text1 = list1.SelectedItem.ListSubItems.Item(1)
            TEXT1_KeyDown 13, 1
        End If
    End If
End If
End Sub

Private Sub M_AST_CAP_GotFocus()
SLT
End Sub

Private Sub M_AST_DEP_GotFocus()
SLT
End Sub

Private Sub M_AST_LFP_GotFocus()
SLT
End Sub

Private Sub M_AST_LFP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    M_AST_LFP = Val(M_AST_LFP)
    M_AST_RMK.SetFocus
ElseIf KeyCode = 38 Then
    M_AST_LFP = Val(M_AST_LFP)
    indidy.SetFocus
End If
End Sub

Private Sub M_AST_LFP_KeYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub M_AST_RMK_Keydown(K As Integer, s As Integer)
If K = 13 Then
    save
End If
End Sub

Private Sub M_AST_TYP_GotFocus()
SLT
End Sub

Private Sub m_ast_typ_Keydown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 38 Then
If CHECKTB("att.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If rs.State = 1 Then rs.Close
    rs.Open " select * from att where att_idy like '" & Sincrement(M_AST_TYP) & "'", con, adOpenStatic
    If rs.RecordCount > 0 Then
        M_AST_TYP = Sincrement(M_AST_TYP)
        M_AST_NME = Trim(rs("ATT_TYP"))
        If KeyCode = 38 Then
            Text2.SetFocus
        Else
            M_CMP_IDY.SetFocus
        End If
    Else
        assettypelist
    End If
End If
End Sub

Private Sub assettypelist()
If CHECKTB("att.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from att where ATT_TYP like '" & Trim(M_AST_TYP) & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add , , "Asset Type"
    list1.ColumnHeaders.Add , , "ID"
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("ATT_TYP")))
        ls.ListSubItems.Add , , Trim(rs("att_idy"))
        rs.MoveNext
    Next
    Label20.Caption = "Select Asset Type"
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "Asset Type Not Found, Define Asset Type !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_AST_TYP.SetFocus
End If
End Sub
Private Sub ASM_LST()
If CHECKTB("asm.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from ASM where ASM_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add , , "Asset Name"
    list1.ColumnHeaders.Add , , "ID"
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("ASM_nme")))
        ls.ListSubItems.Add , , Trim(rs("ASM_idy"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "Asset Not Found, Define Asset !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub M_AST_TYP_keypress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub M_AST_VAL_GotFocus()
SLT
End Sub

Private Sub M_AST_VAL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    M_AST_VAL = convert(M_AST_VAL)
    M_AST_DEP.SetFocus
ElseIf KeyCode = 38 Then
    M_AST_VAL = convert(M_AST_VAL)
    M_AST_CAP.SetFocus
End If
End Sub

Private Sub M_AST_VAL_KEYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_AST_CAP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    M_AST_CAP = QTY(M_AST_CAP)
    M_AST_VAL.SetFocus
ElseIf KeyCode = 38 Then
    M_AST_CAP = QTY(M_AST_CAP)
    M_MNT_IDY.SetFocus
End If
End Sub

Private Sub M_AST_CAP_KEYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_AST_DEP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    M_AST_DEP = convert(M_AST_DEP)
    indidy.SetFocus
ElseIf KeyCode = 38 Then
    M_AST_DEP = convert(M_AST_DEP)
    M_AST_VAL.SetFocus
End If
End Sub

Private Sub M_AST_DEP_KEYPRESS(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_CMP_IDY_GotFocus()
SLT
End Sub

Private Sub M_CMP_IDY_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    M_MFG_IDY.SetFocus
ElseIf K = 38 Then
    M_AST_TYP.SetFocus
End If
End Sub

Private Sub M_CMP_IDY_KeyPRESS(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub M_MFG_IDY_GotFocus()
SLT
End Sub

Private Sub M_mfg_IDY_keydown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If

    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_MFG_IDY)) & "' AND GRP_CLF in('Debtor','Creditor') AND DBT_TYP='Y'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        M_MFG_IDY = Sincrement(M_MFG_IDY)
        M_MFG_NME = Trim(REC_SET("act_nme") & "")
        If K = 38 Then
            M_CMP_IDY.SetFocus
        Else
            M_MNT_IDY.SetFocus
        End If
    Else
        V_DLR_LST
    End If
End If
End Sub

Private Sub M_MFG_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub M_MNT_IDY_GotFocus()
SLT
End Sub

Private Sub M_mnt_IDY_keydown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset, REC_SET1 As New ADODB.Recordset, REC_SET2 As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If

    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(M_MNT_IDY)) & "'  AND GRP_CLF in('Debtor','Creditor') and DBT_TYP='Y'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        M_MNT_IDY = Sincrement(M_MNT_IDY)
        M_MNT_NME = Trim(REC_SET("act_nme") & "")
        If K = 38 Then
            M_MFG_IDY.SetFocus
        Else
            M_AST_CAP.SetFocus
        End If
    Else
        V_MNT_LST
    End If
End If
End Sub
Private Sub V_MNT_LST()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from act where GRP_CLF IN('Debtor','Creditor') and act_nme like '" & Trim(M_MNT_IDY) & "%' AND  DBT_TYP ='Y'", con, adOpenKeyset, adLockPessimistic
list1.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
list1.ColumnHeaders.CLEAR
list1.ColumnHeaders.Add , , "Name"
list1.ColumnHeaders.Add , , "ID"
list1.ColumnHeaders.Add , , "Address"
list1.ColumnHeaders.Item(1).Width = 3500
list1.ColumnHeaders.Item(2).Width = 1500
list1.ColumnHeaders.Item(3).Width = 2000
    For A = 1 To REC_SET.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(REC_SET("ACT_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Maintenance"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Maintenance Not Found, Define Maintenance !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        M_MNT_IDY.SetFocus
    End If
End If
End Sub
Private Sub V_DLR_LST()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from act where GRP_CLF IN('Debtor','Creditor') and act_nme like '" & Trim(M_MFG_IDY) & "%' AND  DBT_TYP ='Y'", con, adOpenKeyset, adLockPessimistic
list1.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
list1.ColumnHeaders.CLEAR
list1.ColumnHeaders.Add , , "Name"
list1.ColumnHeaders.Add , , "ID"
list1.ColumnHeaders.Add , , "Address"
list1.ColumnHeaders.Item(1).Width = 3500
list1.ColumnHeaders.Item(2).Width = 1500
list1.ColumnHeaders.Item(3).Width = 2000
    For A = 1 To REC_SET.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(REC_SET("ACT_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_ad4") & "")
        REC_SET.MoveNext
    Next
    
    SSTab1.Tab = 1
    Label20 = "Select Manufacturer"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Manufacturer Not Found, Define Manufacturer !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        M_MFG_IDY.SetFocus
    End If
End If
End Sub

Private Sub M_MNT_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub SSTab1_DblClick()
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
M_MDI.STS_BAR.Panels(1).Text = "Type Asset  ID or Press <Enter> to get List of Asset  ID's"
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 27 And Text1 = "" Then Unload Me
End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Asset Name"
SLT
End Sub

Private Sub TEXT1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo lab
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
If CHECKTB("asm.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

If rs.State = 1 Then rs.Close
        rs.Open " select * from ASM where ASM_idy like '" & Sincrement(Text1) & "'", con, adOpenStatic
        If rs.RecordCount > 0 Then
            Text1 = Sincrement(Text1)
            Text2 = Trim(rs("ASM_nme"))
            M_AST_TYP = Trim(rs!ATT_IDY)
            M_CMP_IDY = Trim(rs!ASM_CID)
            M_AST_VAL = convert(Val(rs!ASM_VAL))
            M_AST_DEP = convert(Val(rs!ASM_DPV))
            M_MNT_IDY = Trim(rs!ACT_IDY)
            M_MFG_IDY = Trim(rs!MFG_IDY)
            M_AST_CAP = QTY(Val(Trim(rs!ASM_CAP & "")))
            M_AST_RMK = Trim(rs!ASM_TSP)
            indidy = Trim(rs!ind_idy)
            M_AST_LFP = Val(rs!ASM_LFP)
            M_AST_NME = Trim(GetValue("ATT_TYP", "ATT WHERE ATT_IDY LIKE '" & Trim(M_ATT_TYP) & "'"))
            M_MNT_NME = Trim(GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(M_MNT_IDY) & "'"))
            M_MFG_NME = Trim(GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(M_MFG_IDY) & "'"))
            Text1.Enabled = False
            Text2.SetFocus
            Toolbar1.Buttons(12).Enabled = True
            MOD1 = True
        Else
            Assetlist
        End If
End If
Exit Sub
lab:
MsgBox err.Description
End Sub

Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Text2)) > 0 Then
        M_AST_TYP.SetFocus
     Else
        MsgBox "Asset Name Should Not Be Blank, Give Name !", vbExclamation, "DAS Version 4.0"
        Text2.SetFocus
     End If
End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
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
    If M_FRM_ASM.ActiveControl.name = "Text1" Then Assetlist
ElseIf Button.KEY = "d" Then
    delete
ElseIf Button.KEY = "pre" Then
    M_FRM_PRE.Caption = "Asset ID Prefix"
    M_FRM_PRE.Show 1
ElseIf Button.KEY = "p" Then
'    Openreport CURDIR & "\Assets.RP1", "Assets", 0, "", 6
ElseIf Button.KEY = "pr" Then
    'MDIForm1.RRReport1.WindowTitle = "Retailers' Categories List Report"
    'MDIForm1.RRReport1.ReportName = App.Path & "\Channels list.rrw"
    'MDIForm1.RRReport1.Destination = 1
    'MDIForm1.RRReport1.Printer = "?"
    'MDIForm1.RRReport1.RunReport 1
ElseIf Button.KEY = "crt" Then
    CREATE
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
With MDIForm1
    If ButtonMenu.KEY = "AST_CID" Then
        OpenReport CURDIR & "\Assets.RP1", "Assets - Company ID", 0, "", 6
    ElseIf ButtonMenu.KEY = "AST_MFG" Then
        OpenReport CURDIR & "\Assets.RP1", "Assets - Mfg. ID", 0, "", 6
    ElseIf ButtonMenu.KEY = "AST_REC" Then
        OpenReport CURDIR & "\ASSETS.RP1", "Assets", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "AST_POS" Then
        con.Execute "DELETE FROM STL"
        Dim RSSTL As New ADODB.Recordset
        Dim RSAST As New ADODB.Recordset
        RSSTL.Open "SELECT * FROM STL", con, adOpenKeyset, adLockPessimistic
        RSAST.Open "SELECT * FROM AST WHERE TXN_TYP = 'AOU' GROUP BY ASM_COD ORDER BY TXN_DAT DESC", con, adOpenKeyset, adLockPessimistic
        For A = 1 To RSAST.RecordCount
            RSSTL.AddNew
            RSSTL!INV_IDY = Trim(RSAST!ASM_IDY)
            RSSTL!INV_DAT = Trim(RSAST!txn_dat)
            RSSTL!prd_idy = Trim(RSAST!PTY_IDY)
            RSSTL!TXN_RMK = Trim(RSAST!asm_cod)
            RSSTL!TXN_TYP = "AOU"
            RSSTL.Update
            RSAST.MoveNext
        Next
        If RSAST.State = 1 Then RSAST.Close
        RSAST.Open "SELECT * FROM AST WHERE TXN_TYP LIKE 'AIN' AND ASM_COD NOT IN(SELECT ASM_COD FROM AST WHERE TXN_TYP='AOU')", con, adOpenKeyset, adLockPessimistic
        For A = 1 To RSAST.RecordCount
            RSSTL.AddNew
            RSSTL!INV_IDY = Trim(RSAST!ASM_IDY)
            RSSTL!INV_DAT = Trim(RSAST!txn_dat)
            RSSTL!prd_idy = Trim(RSAST!PTY_IDY)
            RSSTL!TXN_RMK = Trim(RSAST!asm_cod)
            RSSTL!TXN_TYP = "AIN"
            RSSTL.Update
            RSAST.MoveNext
        Next
        OpenReport CURDIR & "\ASSETS.RP1", "Assets Position", 0, SAL_QRY, 6
    ElseIf ButtonMenu.KEY = "AST_TRK" Then
        M_REP_ASM.Show 1
    End If
End With
End Sub
