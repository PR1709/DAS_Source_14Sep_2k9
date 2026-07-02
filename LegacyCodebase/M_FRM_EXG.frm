VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form M_FRM_EXG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Day Closure : Secondary Sales Data : Master Data Receive"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "M_FRM_EXG.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleMode       =   0  'User
   ScaleWidth      =   8820
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
            Picture         =   "M_FRM_EXG.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EXG.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EXG.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EXG.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EXG.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EXG.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EXG.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EXG.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EXG.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EXG.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EXG.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_EXG.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5205
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9181
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
      TabCaption(0)   =   "Day Closure Data &Xchange Module"
      TabPicture(0)   =   "M_FRM_EXG.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSComm1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "MAPIMessages1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "MAPISession1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Master Data &Receive"
      TabPicture(1)   =   "M_FRM_EXG.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Command13"
      Tab(1).Control(2)=   "Command12"
      Tab(1).Control(3)=   "Frame5"
      Tab(1).ControlCount=   4
      Begin MSMAPI.MAPISession MAPISession1 
         Left            =   6600
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   0   'False
      End
      Begin MSMAPI.MAPIMessages MAPIMessages1 
         Left            =   5760
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   1
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   0   'False
      End
      Begin VB.Frame Frame6 
         Caption         =   "New "
         Height          =   2895
         Left            =   -69480
         TabIndex        =   34
         Top             =   480
         Width           =   2655
         Begin VB.CheckBox Check11 
            Caption         =   "&MRP"
            Height          =   375
            Left            =   220
            TabIndex        =   10
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Mfg - BBD Batch"
            Height          =   195
            Left            =   480
            TabIndex        =   38
            Top             =   2370
            Width           =   1200
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Tax Structure - VAT"
            Height          =   195
            Left            =   480
            TabIndex        =   37
            Top             =   1890
            Width           =   1410
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Sale Rate"
            Height          =   195
            Left            =   480
            TabIndex        =   36
            Top             =   930
            Width           =   705
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Purchase Rate"
            Height          =   195
            Left            =   480
            TabIndex        =   35
            Top             =   450
            Width           =   1065
         End
      End
      Begin VB.CommandButton Command13 
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   -67945
         TabIndex        =   12
         Top             =   3480
         Width           =   1120
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&OK"
         Height          =   400
         Left            =   -69120
         TabIndex        =   11
         Top             =   3480
         Width           =   1120
      End
      Begin VB.Frame Frame5 
         Caption         =   "Check that apply "
         Height          =   4095
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   5175
         Begin VB.CommandButton Command4 
            Caption         =   "&Apply"
            Height          =   400
            Left            =   3720
            TabIndex        =   43
            Top             =   3240
            Width           =   1120
         End
         Begin VB.TextBox M_CRT_VAT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2520
            TabIndex        =   42
            Text            =   "Text2"
            Top             =   3480
            Width           =   975
         End
         Begin VB.TextBox M_CRT_RNT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2520
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   3000
            Width           =   975
         End
         Begin VB.CheckBox Check10 
            Caption         =   "&9 Distributor Profile"
            Height          =   255
            Left            =   2520
            TabIndex        =   9
            Top             =   1860
            Width           =   1935
         End
         Begin VB.CheckBox Check9 
            Caption         =   "&5 Glass - Shells"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   2340
            Width           =   1575
         End
         Begin VB.CheckBox Check8 
            Caption         =   "&3 Key Accounts"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   1380
            Width           =   1575
         End
         Begin VB.CheckBox Check7 
            Caption         =   "&7 Districts"
            Height          =   255
            Left            =   2520
            TabIndex        =   7
            Top             =   900
            Width           =   1095
         End
         Begin VB.CheckBox Check6 
            Caption         =   "&8 Regions : Units"
            Height          =   255
            Left            =   2520
            TabIndex        =   8
            Top             =   1380
            Width           =   1575
         End
         Begin VB.CheckBox Check5 
            Caption         =   "&6 Sales Calendar"
            Height          =   255
            Left            =   2520
            TabIndex        =   6
            Top             =   420
            Width           =   1815
         End
         Begin VB.CheckBox Check4 
            Caption         =   "&2 Retailer Type"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   900
            Width           =   1695
         End
         Begin VB.CheckBox Check3 
            Caption         =   "&4 SKU - Products"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   1860
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            Caption         =   "&1 Channel"
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   413
            Width           =   1215
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   5160
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "VAT% on Crate - Shell Rental"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   3570
            Width           =   2070
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Crate - Shell Rental"
            Height          =   195
            Left            =   945
            TabIndex        =   39
            Top             =   3090
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Distributor Data Send "
         Height          =   3975
         Left            =   4560
         TabIndex        =   15
         Top             =   960
         Width           =   4335
         Begin VB.CommandButton Command15 
            Caption         =   "&Print Log"
            Height          =   375
            Left            =   2880
            TabIndex        =   31
            Top             =   3480
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&File Mail"
            Height          =   375
            Left            =   1500
            TabIndex        =   28
            Top             =   3480
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Select &All"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Frame Frame4 
            Caption         =   "Days Selected For Data Transfer - Analytics"
            Height          =   2535
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   4095
            Begin MSComctlLib.ListView LVMOD2 
               Height          =   2205
               Left            =   120
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   240
               Width           =   3825
               _ExtentX        =   6747
               _ExtentY        =   3889
               View            =   3
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HotTracking     =   -1  'True
               HoverSelection  =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
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
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Day Closure Dates"
                  Object.Width           =   7056
               EndProperty
            End
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   300
            Left            =   2445
            TabIndex        =   21
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74317827
            CurrentDate     =   38721
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   300
            Left            =   705
            TabIndex        =   23
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74317827
            CurrentDate     =   38721
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "From"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   420
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "To"
            Height          =   195
            Left            =   2145
            TabIndex        =   20
            Top             =   420
            Width           =   195
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Day Closure"
         Height          =   3975
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   4335
         Begin VB.CommandButton Command17 
            Caption         =   "&Generate"
            Height          =   375
            Left            =   2880
            TabIndex        =   33
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton Command16 
            Caption         =   "C&onfirm"
            Height          =   375
            Left            =   1500
            TabIndex        =   32
            Top             =   3480
            Width           =   1335
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Repor&t"
            Height          =   375
            Left            =   2880
            TabIndex        =   30
            Top             =   3480
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Select All"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Caption         =   "Pending Day Closures"
            Height          =   2535
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   4095
            Begin MSComctlLib.ListView LVMOD 
               Height          =   2205
               Left            =   120
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   240
               Width           =   3825
               _ExtentX        =   6747
               _ExtentY        =   3889
               View            =   3
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               AllowReorder    =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HotTracking     =   -1  'True
               HoverSelection  =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
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
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Day Closure Date"
                  Object.Width           =   6174
               EndProperty
            End
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   720
            TabIndex        =   17
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   74317827
            CurrentDate     =   38721
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "As on"
            Height          =   195
            Left            =   135
            TabIndex        =   16
            Top             =   420
            Width           =   405
         End
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   8400
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Key             =   "ed"
            Object.ToolTipText     =   "View (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Area ID"
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
      Begin MSComDlg.CommonDialog cd1 
         Left            =   6000
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "xls"
         DialogTitle     =   "Select Excel File"
         FileName        =   "*.xls"
         Filter          =   "*.xls"
      End
   End
End
Attribute VB_Name = "M_FRM_EXG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub STP_VAL()
M_CRT_RNT = "0.00"
M_CRT_VAT = "0.00"
End Sub
Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Check4.SetFocus
End If
End Sub
Private Sub Check10_Click()
If KeyAscii = 13 Then
    Check11.SetFocus
ElseIf KeyAscii = 38 Then
    Check6.SetFocus
End If
End Sub
Private Sub Check10_KeyDown(KeyCode As Integer, Shift As Integer)
M_CRT_RNT.SetFocus
End Sub
Private Sub Check10_LostFocus()
M_CRT_RNT.SetFocus
End Sub
Private Sub Check11_Click()
If KeyAscii = 13 Then
    Command12.SetFocus
ElseIf KeyAscii = 38 Then
    Check10.SetFocus
End If
End Sub
Private Sub CHECK3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check9.SetFocus
ElseIf KeyAscii = 38 Then
    Check8.SetFocus
End If
End Sub
Private Sub Check4_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check8.SetFocus
ElseIf KeyAscii = 38 Then
    Check1.SetFocus
End If
End Sub
Private Sub Check5_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check7.SetFocus
ElseIf KeyAscii = 38 Then
    Check9.SetFocus
End If
End Sub
Private Sub Check6_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check10.SetFocus
ElseIf KeyAscii = 38 Then
    Check7.SetFocus
End If
End Sub
Private Sub Check7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check6.SetFocus
ElseIf KeyAscii = 38 Then
    Check5.SetFocus
End If
End Sub
Private Sub Check8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check3.SetFocus
ElseIf KeyAscii = 38 Then
    Check4.SetFocus
End If
End Sub
Private Sub Check9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Check5.SetFocus
ElseIf KeyAscii = 38 Then
    Check3.SetFocus
End If
End Sub
Private Sub Command1_Click()
If Command1.Caption = "&Select All" Then
    For i = 1 To LVMOD.ListItems.count
        LVMOD.ListItems(i).Checked = True
        Command1.Enabled = True
    Next
    Command1.Caption = "&Deselect All"
    Command16.SetFocus
Else
    For i = 1 To LVMOD.ListItems.count
        LVMOD.ListItems(i).Checked = False
    Next
    Command1.Caption = "&Select All"
    Command1.SetFocus
End If
End Sub
Private Function ADS(rs As ADODB.Recordset, FLD As String, DB As String)
If rs(FLD).DefinedSize < Len(Trim(DB)) Then
    ADS = Mid(DB, 1, rs(FLD).DefinedSize)
Else
    ADS = Trim(DB)
End If
End Function
Private Sub Command12_Click()

' Data Import-Export Folder Creation
Dim DEX As String
DEX = "C:\DEX"
If Dir(DEX, vbDirectory) = "" Then
   MkDir DEX
End If
' Data Import-Export Folder Creation

If Check1.Value = 1 Then
    CHN_MEX                     ' Channel Data into CAT.DBF
End If
If Check4.Value = 1 Then
    RTT_MEX                     ' Retailer Type into DIR.DBF
End If
If Check5.Value = 1 Then
    KOC_MEX                     ' Sales Calendar into WEK.DBF
End If
If Check6.Value = 1 Then
    REG_MEX                     ' Region Units into DIR.DBF
End If
If Check7.Value = 1 Then
    DST_MEX                     ' Districts States into DIR.DBF
End If
If Check8.Value = 1 Then
    KAC_MEX                     ' Key Accounts into DIR.DBF
End If
If Check10.Value = 1 Then
    DPF_MEX                     ' Distributor Profile into DIR.DBF
End If
If Check3.Value = 1 Then
    PRD_MEX                     ' PGR, PRD into PGR.DBF, PRD.DBF
End If
If Check9.Value = 1 Then
    SHL_MEX                     ' Shells into SHL.DBF
End If
If Check11.Value = 1 Then
    If MsgBox("Update Existing SKU Batches ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
        BAT_MEX                     ' Update Existing Batches in BRD.DBF
    Else
        If MsgBox("Confirm New SKU Batch Creation ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
           N_BAT_MEX                   ' New Batches in BRD.DBF
        End If
    End If
End If
End Sub
Private Sub Command13_Click()
cancel
Unload Me
End Sub
Private Sub Command16_Click()
Dim CHKPNG As String
CHKPNG = "Y"
If LVMOD.ListItems.count > 0 Then
'.ListItems.Item.Checked = True Then
    Dim LSTIND As Integer
    LSTIND = LVMOD.ListItems.count
    For J = 1 To LSTIND
    If LVMOD.ListItems.Item(J).Checked = False Then
        MsgBox ("Day Closures are Pending Before Selected Dates, Complete Earlier Day Closures"), vbCritical, "DAS Version 4.0"
        cancel
        Exit Sub
        Exit For
        DTPicker1.SetFocus
    End If
    Next
End If

'If CHKPNG = "Y" Then
    Dim Pending_RS As New ADODB.Recordset
    If MsgBox("Confirm Day Closures for the above Dates ? Process Irreversible !", vbQuestion + vbOKCancel, "DAS Version 4.0 - Day Closure Confirmation") = vbOK Then
        Dim REC_VAN As New ADODB.Recordset
        If REC_VAN.State = 1 Then REC_VAN.Close
        REC_VAN.Open "SELECT * FROM VAN WHERE INV_DAT <= CTOD('" & DTPicker1 & "')", con, adOpenKeyset, adLockPessimistic
        If REC_VAN.RecordCount > 1 Then
            REC_VAN.MoveFirst
            While REC_VAN.EOF = False
                If Trim(REC_VAN("VAN_CLS")) <> "C" Then
                    MsgBox ("Van Challans are Not Reconciled and Closed! Van Closures Till This Date Required!"), vbCritical, "DAS Version 4.0"
                    cancel
                    Exit Sub
                End If
                REC_VAN.MoveNext
            Wend
        End If
        For i = 1 To LVMOD.ListItems.count
        If LVMOD.ListItems(i).Checked = True Then
            If Pending_RS.State = 1 Then Pending_RS.Close
            Dim CHK_DAT As Date
            CHK_DAT = Trim(LVMOD.ListItems(i))
            CHK_DAT = Format(CHK_DAT, "DD/MM/YYYY")
            Pending_RS.Open "select * from DIR where substr(alltrim(DIR_IDY),1,3) = 'DAY' .and. DIR_DAT = CTOD('" & CHK_DAT & "')", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
            If Pending_RS.RecordCount > 0 Then
               Pending_RS("DIR_MSG") = "Y"
               Pending_RS.Update
            End If
        End If
        Next
        DTPicker4.SetFocus
    Else
        MsgBox ("No Updates Done!"), vbInformation, "DAS Version 4.0"
        DTPicker1.SetFocus
    End If
'Else
'    MsgBox ("No Updates Done!"), vbInformation, "DAS Version 4.0"
'End If
cancel
End Sub

Private Sub Command17_Click()
'If KeyCode = 13 Then
' * Start of : Correction of Earlier DIR_DAT NULL ERROR through this code patch

    Dim RS_DAYCLS As New ADODB.Recordset
    If RS_DAYCLS.State = 1 Then RS_DAYCLS.Close
    RS_DAYCLS.Open "select * from DIR where substr(alltrim(DIR_IDY),1,3) = 'DAY'", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
    While RS_DAYCLS.EOF = False
        If IsNull(RS_DAYCLS("DIR_DAT")) = True Then
            RS_DAYCLS.delete
            RS_DAYCLS.Update
        End If
        RS_DAYCLS.MoveNext
    Wend
    
' * End of : Correction of Earlier DIR_DAT NULL ERROR through this code patch
    LVMOD.ListItems.CLEAR
    Dim rs As New ADODB.Recordset
    Dim DD As Integer
    Dim LDD As Integer
    Dim id As Date
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM YER WHERE ACT_PAS = 'Y'", con, adOpenKeyset
    Dim ST_DAT As Date
    Dim EN_DAT As Date
    Dim BEG_DAT As Date
    Dim END_DAT As Date
    
    ST_DAT = rs("TO_DAT")
    EN_DAT = rs("FRM_DAT")
'    DD = ST_DAT - EN_DAT
    DD = DateDiff("D", rs("FRM_DAT"), rs("TO_DAT"))
    For i = 0 To DD
    '    ID = EN_DAT + I
        id = DateAdd("Y", i, EN_DAT)
        'Dim RS_DAYCLS As New ADODB.Recordset
        If RS_DAYCLS.State = 1 Then RS_DAYCLS.Close
        RS_DAYCLS.Open "select * from DIR where substr(alltrim(DIR_IDY),1,3) = 'DAY' .and. DIR_DAT = CTOD('" & id & "')", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh
        If RS_DAYCLS.RecordCount <= 0 Then
            RS_DAYCLS.AddNew
            RS_DAYCLS("DIR_IDY") = "DAY"
            RS_DAYCLS("DIR_IDX") = ""
            RS_DAYCLS("DIR_DAT") = id
            RS_DAYCLS("DIR_ALS") = id
            RS_DAYCLS("DIR_MSG") = "N"
            RS_DAYCLS.Update
        End If
    Next i
    Dim Pending_RS As New ADODB.Recordset
    If Pending_RS.State = 1 Then Pending_RS.Close
    Pending_RS.Open "select * from DIR where substr(alltrim(DIR_IDY),1,3) = 'DAY' .and. alltrim(DIR_MSG) = 'N'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
        If Pending_RS.RecordCount > 0 Then
            Pending_RS.MoveFirst
'            If IsNull(CDate(Trim(Pending_RS("DIR_DAT")))) = True Then
'                Pending_RS.delete
'                Pending_RS.Update
'            End If
            For i = 1 To Pending_RS.RecordCount
                If CDate(Trim(Pending_RS("DIR_DAT"))) <= DTPicker1.Value Then
                    LVMOD.ListItems.Add = Format(Trim(Pending_RS("DIR_DAT")), "DD/MM/YYYY")
                End If
                Pending_RS.MoveNext
            Next
                LVMOD.SetFocus
                If LVMOD.ListItems.count > 0 Then
                    LVMOD.ListItems(1).Selected = True
                Else
                    MsgBox ("There are no Pending Day Closures As On This Date"), vbInformation, "DAS Version 4.0"
                    cancel
                End If
            Else
                MsgBox ("There are no Pending Day Closures As On This Date"), vbInformation, "DAS Version 4.0"
                cancel
        End If
End Sub
Private Sub Command2_Click()
If Command2.Caption = "Select &All" Then
    For i = 1 To LVMOD2.ListItems.count
        LVMOD2.ListItems(i).Checked = True
        Command2.Enabled = True
    Next
    Command2.Caption = "Deselect A&ll"
    Command3.SetFocus
Else
    For i = 1 To LVMOD2.ListItems.count
        LVMOD2.ListItems(i).Checked = False
    Next
    Command2.Caption = "Select &All"
    Command2.SetFocus
End If
End Sub
Private Sub MM_EX_FILE()
' * Distributor ID Identification
        Dim DIR_RS As New ADODB.Recordset
        Dim M_WDS_IDY As String
        Dim M_WDS_SRL As String
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            If DIR_RS.BOF = False Then
                DIR_RS.MoveFirst
                M_WDS_IDY = Trim(DIR_RS("DIR_ALS"))
            End If
            If Len(M_WDS_IDY) < 8 Then
                MsgBox ("Distributor ID Required for this Process! Obtain Distributor ID and Start Again"), vbCritical, "DAS Version 4.0"
'                cancel
                Exit Sub
            End If
            M_WDS_SRL = (Val(DIR_RS("DIR_MSG")) + 1)
            If Len(Trim(M_WDS_SRL)) < 3 Then
                 While Len(Trim(M_WDS_SRL)) < 3
                      M_WDS_SRL = "0" & M_WDS_SRL
                 Wend
            End If
        End If
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'SYS' and DIR_IDX = 'VER'", con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            If DIR_RS.BOF = False Then
                DIR_RS.MoveFirst
                DIR_RS("DIR_MSG") = Trim("01.01.01")
                DIR_RS.Update
            End If
        End If
' * Distributor ID Identification
        Dim Pending_RS As New ADODB.Recordset
        If Pending_RS.State = 1 Then Pending_RS.Close
        Pending_RS.Open "select * from DIR where DIR_DAT >= CTOD('" & Format(DTPicker4, "DD/MM/YYYY") & "') AND DIR_DAT <= CTOD('" & Format(DTPicker3, "DD/MM/YYYY") & "') AND alltrim(DIR_MSG) = 'Y' AND substr(alltrim(DIR_IDY),1,3) = 'DAY'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
        'SELECT * FROM DIR WHERE CDOW(DIR_DAT) = "Friday" AND TRIM(dir_msg) = "Y"
        If Pending_RS.RecordCount > 0 Then
            Pending_RS.MoveFirst
            ST_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
            Pending_RS.MoveLast
            EN_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
        End If

If MsgBox("Create Data File for Data Transfer Upto " & Format(DTPicker3.Value, "DD/MM/YYYY"), vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
        
    ' * : Beginning of DAS_CUSTOMER.DAT Creation Code for Data Transfer
    Me.MousePointer = vbHourglass
    Dim F_OMS, F_LOC, F_MAR, F_SMN
    Dim FS_SOU, FS_LOC, FS_MAR, FS_SMN As Object
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    Set FS_LOC = CreateObject("Scripting.filesystemobject")
    Set FS_MAR = CreateObject("SCripting.filesystemobject")
    Set FS_SMN = CreateObject("SCripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\DAS_CUSTOMER.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DAS_CUSTOMER.DAT")
    End If
'    If FS_LOC.FileExists(CURDIR & "\HHT\LOCATION.TXT") = True Then
'        FS_LOC.DeleteFile (CURDIR & "\HHT\LOCATION.TXT")
'    End If
'    If FS_MAR.FileExists(CURDIR & "\HHT\MARKET.TXT") = True Then
'        FS_MAR.DeleteFile (CURDIR & "\HHT\MARKET.TXT")
'    End If
'    If FS_SMN.FileExists(CURDIR & "\HHT\SALESMAN.TXT") = True Then
'        FS_SMN.DeleteFile (CURDIR & "\HHT\SALESMAN.TXT")
'    End If
    
    Set F_OMS = FS_SOU.OpenTextFile(CURDIR & "\HHT\DAS_CUSTOMER.DAT", 2, True)
'    Set F_LOC = FS_LOC.OpenTextFile(CURDIR & "\HHT\LOCATION.TXT", 2, True)
'    Set F_MAR = FS_MAR.OpenTextFile(CURDIR & "\HHT\MARKET.TXT", 2, True)
'    Set F_SMN = FS_SMN.OpenTextFile(CURDIR & "\HHT\SALESMAN.TXT", 2, True)
    
    Dim J As Integer
    J = 0
    
    
    Dim Category_RS As New ADODB.Recordset
    Dim Location_RS As New ADODB.Recordset
    Dim Market_RS As New ADODB.Recordset
    Dim Salesman_RS As New ADODB.Recordset
    Dim Invoice_Chk_RS As New ADODB.Recordset
    
    Dim Retailer_RS As New ADODB.Recordset
    If Retailer_RS.State = 1 Then Retailer_RS.Close
    Retailer_RS.Open "select * from ACT where ACT_TYP = 'SAL' or ACT_TYP = 'PUR'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh  CST_TYP <> 'Van'
    
    If Retailer_RS.RecordCount > 0 Then Retailer_RS.MoveFirst
       
       Dim L_CAT_NME As String
       Dim L_LOC_NME As String
       Dim L_MAR_NME As String
       Dim L_DST_NME As String
       Dim L_MAR_GRD As String
       
       If Category_RS.State = 1 Then Category_RS.Close
       Category_RS.Open "SELECT * FROM CAT", con, adOpenKeyset, adLockPessimistic
       If Location_RS.State = 1 Then Location_RS.Close
       Location_RS.Open "SELECT * FROM LOC", con, adOpenKeyset, adLockPessimistic
       If Market_RS.State = 1 Then Market_RS.Close
       Market_RS.Open "SELECT * FROM MAR", con, adOpenKeyset, adLockPessimistic
       
       While Not Retailer_RS.EOF
       
          If Category_RS.RecordCount > 0 Then Category_RS.MoveFirst
          Category_RS.Find "CAT_IDY='" & Trim(Retailer_RS("CAT_IDY")) & "'"
          If Category_RS.EOF = False Then
             L_CAT_NME = Category_RS("CAT_NME")
          End If
          
          If Location_RS.RecordCount > 0 Then Location_RS.MoveFirst
          Location_RS.Find "LOC_IDY='" & Trim(Retailer_RS("LOC_IDY")) & "'"
          If Location_RS.EOF = False Then
             L_LOC_NME = Location_RS("LOC_NME")
          End If
          
          If Market_RS.RecordCount > 0 Then Market_RS.MoveFirst
          Market_RS.Find "MAR_IDY='" & Trim(Retailer_RS("MAR_IDY")) & "'"
          If Market_RS.EOF = False Then
             L_MAR_NME = Market_RS("MAR_NME")
             L_DST_NME = Market_RS("MAR_PGR")
             L_MAR_GRD = Market_RS("MAR_GRD")
             L_MAR_CSG = Market_RS("MAR_CSG")
             L_MAR_CLU = Market_RS("MAR_CLU")
             L_MAR_ROC = Market_RS("MAR_ROC")
             L_MAR_TWC = Market_RS("MAR_TWC")
             L_MAR_TWN = Market_RS("MAR_TWN")
             L_MAR_DST = Market_RS("MAR_DST")
             L_MAR_STA = Market_RS("MAR_STA")
             L_MAR_COU = Market_RS("MAR_COU")
          End If
          
'          F_OMS.WriteLine Trim(M_WDS_IDY) & "|" & Trim(Retailer_RS("ACT_IDY")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(L_DST_NME) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(L_LOC_NME) & "|" & Trim(L_CAT_NME) & "|" & Trim(L_MAR_NME) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("DBT_TYP"))
          If Trim(Retailer_RS("ACT_TYP")) = "PUR" Then
                'F_OMS.WriteLine Trim(Retailer_RS("ACT_IDY")) & "|" & "S" & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(Retailer_RS("LOC_IDY")) & "|" & Trim(Retailer_RS("MAR_IDY")) & "|" & Trim(Retailer_RS("CAT_IDY")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("DBT_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("ROU_IDY"))
                'J = J + 1
          Else
                'F_OMS.WriteLine Trim(Retailer_RS("ACT_IDY")) & "|" & "C" & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(Retailer_RS("LOC_IDY")) & "|" & Trim(Retailer_RS("MAR_IDY")) & "|" & Trim(Retailer_RS("CAT_IDY")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("DBT_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("ROU_IDY"))
                If Invoice_Chk_RS.State = 1 Then Invoice_Chk_RS.Close
                Invoice_Chk_RS.Open "SELECT * FROM INL WHERE CST_IDY = '" & Trim(Retailer_RS("ACT_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If Invoice_Chk_RS.RecordCount > 0 Then
                    F_OMS.WriteLine Trim(M_WDS_IDY) & Right(Trim(Retailer_RS("ACT_IDY")), 4) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & "I" & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_ZIP")) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_FAX")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(L_MAR_COU) & "|" & Trim(L_MAR_STA) & "|" & Trim(L_MAR_DST) & "|" & Trim(L_MAR_TWN) & "|" & Trim(L_CAT_NME) & "|" & Trim(L_MAR_CSG) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(L_MAR_CLU) & "|" & Trim(L_MAR_ROC) & "|" & Trim(L_MAR_TWC) & "|" & Trim(Retailer_RS("DBT_TYP"))
                    J = J + 1
                End If
          End If
          Retailer_RS.MoveNext
       Wend
'''       MsgBox "" & J & ": Outlet Records Created for Data Transfer!"
'''    Else
'''       MsgBox ("No Outlet Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
'''       cancel
'''    End If
'       Dim n As Integer
'       n = 0
'       If Location_RS.State = 1 Then Location_RS.Close
'       Location_RS.Open "SELECT * FROM LOC", CON, adOpenKeyset, adlockpessimistic
'       While Not Location_RS.EOF
'          F_LOC.WriteLine Trim(Location_RS("LOC_IDY")) & "|" & Trim(Location_RS("LOC_NME"))
'          Location_RS.MoveNext
'          n = n + 1
'       Wend
'
'       Dim o As Integer
'       o = 0
'       If Market_RS.State = 1 Then Market_RS.Close
'       Market_RS.Open "SELECT * FROM MAR", CON, adOpenKeyset, adlockpessimistic
'       While Not Market_RS.EOF
'          F_MAR.WriteLine Trim(Market_RS("MAR_IDY")) & "|" & Trim(Market_RS("MAR_NME")) & "|" & Trim(Market_RS("MAR_PGR")) & "|" & Trim(Market_RS("MAR_GRD"))
'          Market_RS.MoveNext
'          o = o + 1
'       Wend
'
'       Dim P As Integer
'       P = 0
'       If Salesman_RS.State = 1 Then Salesman_RS.Close
'       Salesman_RS.Open "SELECT * FROM SMN", CON, adOpenKeyset, adlockpessimistic
'       While Not Salesman_RS.EOF()
'          F_SMN.WriteLine Trim(Salesman_RS("SMN_IDY")) & "|" & Trim(Salesman_RS("SMN_NME"))
'          Salesman_RS.MoveNext
'          P = P + 1
'       Wend
'
        F_OMS.Close
'        F_LOC.Close
'        F_MAR.Close
'        F_SMN.Close
'    cancel
    
    ' * : End of DAS_CUSTOMER.DAT Creation Code for Data Transfer
    
    ' * : Beginning of DAS_SECONDARY.DAT Creation Code for Data Transfer
    
    Dim F_INV
    Dim F_INL
    Dim F_VAN
    Dim F_EIS
    
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    Set FS_INL = CreateObject("Scripting.filesystemobject")
'    Set FS_VAN = CreateObject("Scripting.filesystemobject")
'    Set FS_EIS = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\DAS_SECONDARY.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DAS_SECONDARY.DAT")
    End If
    
    If FS_INL.FileExists(CURDIR & "\HHT\DAS_PURCHASE.DAT") = True Then
        FS_INL.DeleteFile (CURDIR & "\HHT\DAS_PURCHASE.DAT")
    End If
    
'    If FS_VAN.FileExists(CURDIR & "\HHT\LOAD_SHEET.TXT") = True Then
'        FS_VAN.DeleteFile (CURDIR & "\HHT\LOAD_SHEET.TXT")
'    End If
'
'    If FS_EIS.FileExists(CURDIR & "\HHT\EMPTY_ISSUE.TXT") = True Then
'        FS_EIS.DeleteFile (CURDIR & "\HHT\EMPTY_ISSUE.TXT")
'    End If
    
    Set F_INV = FS_SOU.OpenTextFile(CURDIR & "\HHT\DAS_SECONDARY.DAT", 2, True)
    Set F_INL = FS_INL.OpenTextFile(CURDIR & "\HHT\DAS_PURCHASE.DAT", 2, True)
'    Set F_VAN = FS_VAN.OpenTextFile(CURDIR & "\HHT\LOAD_SHEET.TXT", 2, True)
'    Set F_EIS = FS_EIS.OpenTextFile(CURDIR & "\HHT\EMPTY_ISSUE.TXT", 2, True)
    
    Dim K As Double
    Dim r As Double
    Dim s As Double
    Dim T As Double
    K = 0
    r = 0
    s = 0
    T = 0
    Dim INL_RS As New ADODB.Recordset
    Dim INV_RS As New ADODB.Recordset
    Dim LNV_RS As New ADODB.Recordset
    Dim PRD_RS As New ADODB.Recordset
    Dim BAT_RS As New ADODB.Recordset
    Dim SCH_RS As New ADODB.Recordset
    Dim VAN_RS As New ADODB.Recordset
    Dim SNV_RS As New ADODB.Recordset
    Dim ACT_RS As New ADODB.Recordset
    
    Dim L_CST_IDY As String
    Dim L_ORD_IDY As String
    Dim L_BAT_NME As String
    Dim L_PRD_COD As String
    Dim L_MER_NUM As String
    
    If INL_RS.State = 1 Then INL_RS.Close
    If INV_RS.State = 1 Then INV_RS.Close
    If LNV_RS.State = 1 Then LNV_RS.Close
    If PRD_RS.State = 1 Then PRD_RS.Close
    If BAT_RS.State = 1 Then BAT_RS.Close
    If SCH_RS.State = 1 Then SCH_RS.Close
       
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * FROM DIR where substr(alltrim(DIR_IDY),1,3) = 'DAY' AND alltrim(DIR_MSG) = 'Y' AND DIR_DAT >= CTOD('" & DTPicker4 & "') AND DIR_DAT <= CTOD('" & DTPicker3 & "')", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveFirst
        ST_DAT = DIR_RS("DIR_DAT")
        BEG_DAT = DIR_RS("DIR_DAT")
        DIR_RS.MoveLast
        EN_DAT = DIR_RS("DIR_DAT")
        END_DAT = DIR_RS("DIR_DAT")
    End If
    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    INV_RS.Open "SELECT * FROM INV WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    BAT_RS.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
    SCH_RS.Open "SELECT * FROM SCH", con, adOpenKeyset, adLockPessimistic
    
    If INV_RS.RecordCount > 0 Then INV_RS.MoveFirst
    
        While Not INV_RS.EOF

            If PRD_RS.RecordCount > 0 Then PRD_RS.MoveFirst
            PRD_RS.Find "PRD_IDY='" & Trim(INV_RS("PRD_IDY")) & "'"
            If PRD_RS.EOF = False Then
               L_PRD_COD = PRD_RS("ALT_IDY")
            End If
            
            If BAT_RS.RecordCount > 0 Then BAT_RS.MoveFirst
            BAT_RS.Find "BAT_IDY='" & Trim(INV_RS("BAT_IDY")) & "'"
            If BAT_RS.EOF = False Then
               L_BAT_NME = BAT_RS("BAT_NME")
               If IsNull(BAT_RS("U_PUR_LPP")) = True Then
                    L_PUR_LPP = 0#
               Else
                    L_PUR_LPP = Val(BAT_RS("U_PUR_LPP"))
               End If
               If IsNull(BAT_RS("U_SAL_LPP")) = True Then
                    L_SAL_LPP = 0#
                    L_RMM_UNT = 0#
               Else
                    L_SAL_LPP = Val(BAT_RS("U_SAL_LPP"))
                    L_RMM_UNT = (Val(BAT_RS("PRD_MRP")) / Val(BAT_RS("SKU_SBU"))) - Val(BAT_RS("U_SAL_LPP"))
               End If
               L_MRP_RAT = (Val(BAT_RS("PRD_MRP")) / Val(BAT_RS("SKU_SBU")))
               If IsNull(BAT_RS("SAL_LPP")) = True Or IsNull(BAT_RS("PUR_LPP")) = True Then
                    L_DMM_UNT = 0
               Else
                    L_DMM_UNT = Val(BAT_RS("SAL_LPP")) - Val(BAT_RS("PUR_LPP")) 'VALUE TAKEN ON PER CASE BASIS - PNR
               End If
               If IsNull(BAT_RS("P_TOT_VAT")) = True Or IsNull(BAT_RS("S_TOT_VAT")) = True Then
                    PL_TOT_VAT = 0#
                    SL_TOT_VAT = 0#
               Else
                    PL_TOT_VAT = (Val(BAT_RS("P_TOT_VAT")) / Val(BAT_RS("SKU_SBU")))
                    SL_TOT_VAT = (Val(BAT_RS("S_TOT_VAT")) / Val(BAT_RS("SKU_SBU")))
               End If
            End If
            
            If INL_RS.RecordCount > 0 Then INL_RS.MoveFirst
            INL_RS.Find "INV_IDY='" & Trim(INV_RS("INV_IDY")) & "'"
            If INL_RS.EOF = False Then
               L_CST_IDY = INL_RS("CST_IDY")
               L_ORD_IDY = INL_RS("ORD_IDY")
               L_LRN_IDY = INL_RS("LRN_IDY")
               L_LRN_DAT = INL_RS("LRN_DAT")
            End If
            
            If SCH_RS.RecordCount > 0 Then SCH_RS.MoveFirst
            SCH_RS.Find "SCH_IDY='" & Trim(INV_RS("SCH_IDY")) & "'"
            If SCH_RS.EOF = False Then
                L_MER_NUM = SCH_RS("MER_NUM")
            End If
            If BAT_RS.EOF = False Then
                'M_WDS_IDY & "|"Trim(STR(Inv_RS("COL_AMT"))) & "|" &
                If Trim(INV_RS("TXN_TYP")) = "SAL" Then
                    If Trim(INV_RS("STK_TYP")) = "Salable" Then
                        'F_INV.WriteLine Trim(INV_RS("INV_IDY")) & "|" & Format(INV_RS("INV_DAT"), "DD/MM/YYYY") & "|" & Trim(INV_RS("TXN_TYP")) & "|" & Trim(L_CST_IDY) & "|" & Trim(L_PRD_COD) & "|" & QTY(INV_RS("PRD_QTY")) & "|" & AMT(INV_RS("PRD_RAT")) & "|" & Trim(L_BAT_NME) & "|" & Format(BAT_RS("BRD_DOM"), "DD/MM/YYYY") & "|" & AMT(BAT_RS("PRD_MRP")) & "|" & Trim(INV_RS("STK_TYP")) & "|" & Trim(L_MER_NUM) & "|" & Trim(L_LRN_IDY) & "|" & Format(L_LRN_DAT, "DD/MM/YYYY") & "|" & Trim(INL_RS("TXN_TPT"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIV" & "|" & Val(INV_RS("TOT_PCS")) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIR" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & (Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & (Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) & "|" & Trim(INV_RS("INV_IDY"))
                        K = K + 5
                    End If
                    If Trim(INV_RS("STK_TYP")) = "Free" Then
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIV" & "|" & Val(INV_RS("TOT_PCS")) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIR" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & (Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & (Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "OMSDA" & "|" & (Val(INV_RS("TOT_PCS")) * Val(L_SAL_LPP)) & "|" & Trim(INV_RS("INV_IDY"))
                        K = K + 6
                        
                        If LNV_RS.State = 1 Then LNV_RS.Close
                        LNV_RS.Open "SELECT * FROM LNV WHERE INV_IDY = '" & Trim(INV_RS("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                        If LNV_RS.RecordCount > 0 Then
                            While Not LNV_RS.EOF
                                If Val(LNV_RS("LNV_AMT")) > 0 Then
                                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "OMSDA" & "|" & Val(LNV_RS("LNV_AMT")) & "|" & Trim(INV_RS("INV_IDY"))
                                    K = K + 1
                                End If
                                LNV_RS.MoveNext
                            Wend
                        End If
                    End If
                End If
                If Trim(INV_RS("TXN_TYP")) = "SRN" Then
                    If Trim(INV_RS("STK_TYP")) = "Salable" Then
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRV" & "|" & (Val(INV_RS("TOT_PCS")) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRR" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        K = K + 5
                    End If
                    If Trim(INV_RS("STK_TYP")) = "Free" Then
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRV" & "|" & (Val(INV_RS("TOT_PCS")) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRR" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "OMSDA" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(L_SAL_LPP)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        K = K + 6
                    End If
                End If
            End If
            INV_RS.MoveNext
        Wend
'        MsgBox "" & K & ": Stock Records Created for Data Transfer!"
'    Else
'       MsgBox ("No Stock Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
'       cancel
'    End If
    F_INV.Close
    On Error Resume Next
    If INL_RS.State = 1 Then INL_RS.Close
    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    If INL_RS.RecordCount > 0 Then INL_RS.MoveFirst
    While Not INL_RS.EOF
        M_ADD_AMT = 0
        M_DIS_AMT = 0
        M_TAX_AMT = 0
        T_SKU_QTY = 0
        If INV_RS.State = 1 Then INV_RS.Close
        INV_RS.Open "SELECT * FROM INV WHERE INV.INV_IDY = '" & Trim(INL_RS("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
        If INV_RS.RecordCount > 0 Then INV_RS.MoveFirst
        While Not INV_RS.EOF
            If INV_RS("PRD_QTY") > 0 Then
                If IsNull(INV_RS("ADD_AMT")) = False Then
                    M_ADD_AMT = M_ADD_AMT + INV_RS("ADD_AMT")
                End If
                If IsNull(INV_RS("DIS_AMT")) = False Then
                    M_DIS_AMT = M_DIS_AMT + INV_RS("DIS_AMT")
                End If
                If IsNull(INV_RS("TAX_AMT")) = False Then
                    M_TAX_AMT = M_TAX_AMT + INV_RS("TAX_AMT")
                End If
                If IsNull(INV_RS("TOT_PCS")) = False Then
                    T_SKU_QTY = Val(T_SKU_QTY) + Val(INV_RS("TOT_PCS"))
                End If
            End If
            INV_RS.MoveNext
        Wend
        If Trim(INL_RS("TXN_TYP")) <> "OPS" Then
            If INV_RS.RecordCount > 0 Then
                If Trim(INL_RS("TXN_TYP")) = "PUR" Then
                    'F_INL.WriteLine Trim(INL_RS("INV_IDY")) & "|" & Format(INL_RS("INV_DAT"), "DD/MM/YYYY") & "|" & INL_RS("TXN_TYP") & "|" & INL_RS("CST_IDY") & "|" & AMT(INL_RS("TAX_AMT") + INL_RS("EXD_AMT") + M_TAX_AMT) & "|" & AMT(INL_RS("SPD_AMT") + INL_RS("CSD_AMT") + M_ADD_AMT + M_DIS_AMT) & "|" & AMT(INL_RS("NET_AMT")) & "|" & Trim(INL_RS("SMN_IDY")) & "|" & Trim(INL_RS("ROU_IDY")) & "|" & Trim(INL_RS("ORD_IDY"))
                    F_INL.WriteLine Format(INL_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(M_WDS_IDY) & "|" & Trim(INL_RS("INV_IDY")) & "|" & Trim(INL_RS("ORD_IDY")) & "|" & Format(INL_RS("DCH_DAT"), "YYYYMMDD") & "|" & Val(T_SKU_QTY) & "|" & Val(INL_RS("NET_AMT"))
                    r = r + 1
                Else
                    'F_INL.WriteLine Trim(INL_RS("INV_IDY")) & "|" & Format(INL_RS("INV_DAT"), "DD/MM/YYYY") & "|" & INL_RS("TXN_TYP") & "|" & INL_RS("CST_IDY") & "|" & AMT(INL_RS("TAX_AMT") + INL_RS("EXD_AMT") + M_TAX_AMT) & "|" & AMT(INL_RS("SPD_AMT") + INL_RS("CSD_AMT") + M_ADD_AMT + M_DIS_AMT) & "|" & AMT(INL_RS("NET_AMT")) & "|" & Trim(INL_RS("SMN_IDY")) & "|" & Trim(INL_RS("ROU_IDY")) & "|" & ""
                    'r = r + 1
                End If
            End If
        End If
        INL_RS.MoveNext
    Wend
    F_INL.Close
    
'    cancel

    ' * : Beginning of DAS_STOCK.DAT Creation Code for Data Transfer

    Dim F_STK
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\DAS_STOCK.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DAS_STOCK.DAT")
    End If
    
    Set F_STK = FS_SOU.OpenTextFile(CURDIR & "\HHT\DAS_STOCK.DAT", 2, True)
    Dim l As Integer
    l = 0
    LDD = DateDiff("D", BEG_DAT, END_DAT)
    LDD = LDD + 1
    Do While LDD <> 0
            STKTXTDEX (BEG_DAT)
            Dim CLD_RS As New ADODB.Recordset
            
            If CLD_RS.State = 1 Then CLD_RS.Close
            If PRD_RS.State = 1 Then PRD_RS.Close
            If BAT_RS.State = 1 Then BAT_RS.Close
            CLD_RS.Open "SELECT PRD.PRD_IDY, SUM(CLS_STK) FROM CLD, BRD, PRD WHERE CLD.PRD_IDY = BRD.BAT_IDY AND BRD.PRD_IDY = PRD.PRD_IDY AND CLD.CLS_STK > 0 ORDER BY PRD.PRD_IDY GROUP BY PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
            PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
            BAT_RS.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
            
            
            If CLD_RS.RecordCount > 0 Then CLD_RS.MoveFirst
            
                While Not CLD_RS.EOF
                               
                    If BAT_RS.RecordCount > 0 Then BAT_RS.MoveFirst
                    BAT_RS.Find "BAT_IDY='" & Trim(CLD_RS("PRD_IDY")) & "'"
                    If BAT_RS.EOF = False Then
                       L_BAT_NME = BAT_RS("BAT_NME")
                    End If
                    
                    If PRD_RS.State = 1 Then PRD_RS.Close
                    PRD_RS.Open "SELECT * FROM PRD WHERE PRD.PRD_IDY = '" & Trim(CLD_RS(0)) & "'", con, adOpenKeyset, adLockPessimistic
                    'If PRD_RS.RecordCount > 0 Then PRD_RS.MoveFirst
                    'PRD_RS.Find "PRD_IDY='" & Trim(BAT_RS("PRD_IDY")) & "'"
                    If PRD_RS.EOF = False Then
                       L_PRD_COD = PRD_RS("ALT_IDY")
                    End If
                                
        '            F_STK.WriteLine Trim(M_WDS_IDY) & "|" & Trim(L_PRD_COD) & "|" & Trim(STR(CLD_RS("OPN_STK"))) & "|" & Trim(STR(CLD_RS("PUR_STK"))) & "|" & Trim(STR(CLD_RS("PRT_STK"))) & "|" & Trim(STR(CLD_RS("SAL_STK"))) & "|" & Trim(STR(CLD_RS("SRN_STK"))) & "|" & Trim(STR(CLD_RS("STI_STK"))) & "|" & Trim(STR(CLD_RS("STO_STK"))) & "|" & Trim(STR(CLD_RS("CLS_STK"))) & "|" & Trim(STR(BAT_RS("PRD_MRP"))) & "|" & Trim(BAT_RS("BAT_NME")) & "|" & BAT_RS("BRD_DOM")
                    'F_STK.WriteLine Format(BEG_DAT, "DD/MM/YYYY") & "|" & Trim(L_PRD_COD) & "|" & Trim(L_BAT_NME) & "|" & BAT_RS("BRD_DOM") & "|" & AMT(BAT_RS("PRD_MRP")) & "|" & QTY(STM(CLD_RS("OPN_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("PUR_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("PRT_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("SAL_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("SRN_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("STI_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("STO_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("CLS_STK"), PRD_RS("SUB_UNT")))
                    F_STK.WriteLine Format(BEG_DAT, "YYYYMMDD") & "|" & Format(END_DAT, "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & "|" & Val(CLD_RS(1))
                    '& "|" & amt(Val(CLD_RS(1)) / Val(PRD_RS("SUB_UNT")))
                    CLD_RS.MoveNext
                    l = l + 1
                Wend
        '        MsgBox "" & L & ": SKU Records Created for Data Transfer!"
        '    Else
        '       MsgBox ("No Stock Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
        '       cancel
        '    End If
           
        '    cancel
            ' * : End of DAS_STOCK.DAT Creation Code for Data Transfer
        BEG_DAT = (BEG_DAT + 1)
        LDD = LDD - 1
    Loop
    F_STK.Close
    
    ' * : Log Text File Creation Begin Process
    
    Dim F_VER
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    If FS_SOU.FileExists(CURDIR & "\HHT\DAS_VERSION.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DAS_VERSION.DAT")
    End If
    
    Dim Q As Integer
    Dim LDC_DAT As Date
    Q = 0
    Set F_VER = FS_SOU.OpenTextFile(CURDIR & "\HHT\DAS_VERSION.DAT", 2, True)
    
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * FROM DIR where substr(alltrim(DIR_IDY),1,3) = 'DAY' AND alltrim(DIR_MSG) = 'Y'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveLast
        LDC_DAT = DIR_RS("DIR_DAT")
    Else
        LDC_DAT = Date
    End If
    M_EXE_VER = Format(FileDateTime(App.Path & "\EFFMCG.EXE"), "dd/mmm/yyyy hh:mm")
    F_VER.WriteLine Trim(M_WDS_IDY) & "|" & Trim("01.01.01") & "|" & Format(DTPicker4, "dd/MM/yyyy") & "|" & Format(DTPicker3, "dd/MM/yyyy") & "|" & Format(LDC_DAT, "dd/MM/yyyy") & "|" & Trim(M_EXE_VER)
    Q = Q + 1
    
    F_VER.Close
    
    Dim F_LOG
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\DAS_LOG.DAT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\DAS_LOG.DAT")
    End If
    
    Set F_LOG = FS_SOU.OpenTextFile(CURDIR & "\HHT\DAS_LOG.DAT", 2, True)
    
    F_LOG.WriteLine "DAS_CUSTOMER.DAT" & "|" & str(J)
    F_LOG.WriteLine "DAS_SECONDARY.DAT" & "|" & str(K)
    F_LOG.WriteLine "DAS_PURCHASE.DAT" & "|" & str(r)
    F_LOG.WriteLine "DAS_STOCK.DAT" & "|" & str(l)
    F_LOG.WriteLine "DAS_VERSION.DAT" & "|" & str(Q)
    
    Dim MDD As String
    Dim MMM As String
    Dim MYY As String
    MDD = Trim(str(Day(DTPicker3)))
    MMM = Trim(str(Month(DTPicker3)))
    MYY = Trim(str(Year(DTPicker3)))
'    F_LOG.WriteLine MDD & MMM & MYY & M_WDS_SRL 'Refers to Day Closure Date, Last Line in Log File is Blocked - PNR
    F_LOG.Close
    ' * : Log Text File Creation End Process
    Me.MousePointer = vbDefault
    MsgBox ("File Creation Process Successful! Log File Generated!"), vbInformation, "DAS Version 4.0"
    
' Beginning of Distributor Information for MAPI Data Xfer
    
    Dim M_NME, M_ADR, M_TFX, M_WDS As String
    Dim RST_DIR As New ADODB.Recordset
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY = 'SYS' .AND. DIR_IDX = 'TT1'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_NME = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY = 'SYS' .AND. DIR_IDX = 'TT2'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_ADR = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY = 'SYS' .AND. DIR_IDX = 'TT3'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_TFX = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY = 'WDS' .AND. DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_WDS = Trim(RST_DIR("DIR_ALS"))
    End If

    If Dir(CURDIR & "\DST_INF.DAT") = "" Then
        Set F_VER = FS_SOU.OpenTextFile(CURDIR & "\DST_INF.DAT", 2, True)
        F_VER.WriteLine "Distributor : " & M_NME
        F_VER.WriteLine "Address     : " & M_ADR
        F_VER.WriteLine "Telephone   : " & M_TFX
        F_VER.WriteLine "DAS Code-ID : " & M_WDS
        F_VER.Close
    End If

' End of Distributor Information for MAPI Data Xfer

    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then DIR_RS.MoveFirst
       DIR_RS("DIR_MSG") = Trim(M_WDS_SRL)
       DIR_RS.Update
    If DIR_RS.State = 1 Then DIR_RS.Close
    'Zip File Creation
            Dim DEX As String
            Dim CDO As String
            DEX = "C:\DEX"
            CDO = "C:\DEX\OutBox"
            If Dir(DEX, vbDirectory) = "" Then
                MkDir DEX
            End If
            If Dir(CDO, vbDirectory) = "" Then
                MkDir CDO
            End If
            
'            CopyFile CURDIR & "\HHT" & "\DAS_CUSTOMER.DAT", CDO & "\DAS_CUSTOMER.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_SECONDARY.DAT", CDO & "\DAS_SECONDARY.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_PURCHASE.DAT", CDO & "\DAS_PURCHASE.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_STOCK.DAT", CDO & "\DAS_STOCK.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_VERSION.DAT", CDO & "\DAS_VERSION.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_LOG.DAT", CDO & "\DAS_LOG.DAT", 0
'            CopyFile CURDIR & "\DST_INF.DAT", CDO & "\DST_INF.DAT", 0
            
            CopyFile CURDIR & "\HHT" & "\DAS_CUSTOMER.DAT", DEX & "\DAS_CUSTOMER.DAT", 0
            CopyFile CURDIR & "\HHT" & "\DAS_SECONDARY.DAT", DEX & "\DAS_SECONDARY.DAT", 0
            CopyFile CURDIR & "\HHT" & "\DAS_PURCHASE.DAT", DEX & "\DAS_PURCHASE.DAT", 0
            CopyFile CURDIR & "\HHT" & "\DAS_STOCK.DAT", DEX & "\DAS_STOCK.DAT", 0
            CopyFile CURDIR & "\HHT" & "\DAS_VERSION.DAT", DEX & "\DAS_VERSION.DAT", 0
            CopyFile CURDIR & "\HHT" & "\DAS_LOG.DAT", DEX & "\DAS_LOG.DAT", 0
            CopyFile CURDIR & "\DST_INF.DAT", DEX & "\DST_INF.DAT", 0
                        
                        
            
            CopyFile App.Path & "\" & "Zip32.dll", CURDIR & "\" & "Zip32.dll", 0
            M_EML_IDY = GetValue("DIR_MSG", "DIR where DIR_IDY = 'EML' AND DIR_IDX = 'IDY'")
            
            M_FIL_NME = "\" & "DAS" & Trim(M_WDS_IDY) & Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & ".ZIP"
            Dim Myzip As ColaZip.GSColaZip32
            Set Myzip = New ColaZip.GSColaZip32
            Myzip.RootDirectory = "\"
            
            Myzip.ZipFileName = "C:\DEX" & M_FIL_NME
            Myzip.AddFile ("C:\DEX" & "\DAS_CUSTOMER.DAT")
            Myzip.AddFile ("C:\DEX" & "\DAS_SECONDARY.DAT")
            Myzip.AddFile ("C:\DEX" & "\DAS_PURCHASE.DAT")
            Myzip.AddFile ("C:\DEX" & "\DAS_STOCK.DAT")
            Myzip.AddFile ("C:\DEX" & "\DAS_VERSION.DAT")
            Myzip.AddFile ("C:\DEX" & "\DAS_LOG.DAT")
            Myzip.AddFile ("C:\DEX" & "\DST_INF.DAT")
            
            Myzip.MakeColaZip ("DAS")
            
            MsgBox ("Zip File " & M_FIL_NME & " Created in the Folder " & DEX), vbInformation, "DAS Version 4.0"
                      
            M_MSG_STR = "Margin Minder Zipped Data From : " & Trim(M_NME) & " , " & Trim(M_ADR) & " , " & Trim(M_TFX) & " , " & Trim(M_WDS) & " , Date: " & Format(Date, ("YYYY/MM/DD")) & " , Time: " & Format(Time, "hh:mm:ss")
            M_MSG_EML = GetValue("DIR_MSG", "DIR WHERE DIR_IDY = 'EML' AND DIR_IDX = 'IDY'")
            M_MSG_SUB = "Zip Data Xfer Margin Minder Data Extract : " & M_WDS & " , Date: " & Format(Date, ("YYYY/MM/DD")) & " , Time: " & Format(Time, "hh:mm:ss")
            
            M_MAIL_ATTMNT = "C:\DEX" & M_FIL_NME
'            If MsgBox("Send Zip File Data to DAS MIS Analytics Desk Now ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
'                If Trim(M_EML_IDY) = "" Then
'                    MailSend "effmcg@consoftsystems.com", Trim(M_MSG_SUB), Trim(M_MSG_STR), Trim(M_MAIL_ATTMNT)
'                    MsgBox ("Data Send via MAPI Executed! Email ID: effmcg@consoftsystems.com" & Trim(M_EML_IDY)), vbInformation, "DAS Version 4.0"
'                Else
'                    MailSend Trim(M_EML_IDY), Trim(M_MSG_SUB), Trim(M_MSG_STR), Trim(M_MAIL_ATTMNT)
'                    MsgBox ("Data Send via MAPI Executed! Email ID: " & Trim(M_EML_IDY)), vbInformation, "DAS Version 4.0"
'                End If
'            Else
'                cancel
'            End If
            CopyFile DEX & M_FIL_NME, CDO & M_FIL_NME, 0
    ' Zip File Creation
            
    cancel
Else
    DTPicker4.SetFocus
End If
End Sub
Private Sub DEXFILE()
' * Distributor ID Identification
        Dim DIR_RS As New ADODB.Recordset
        Dim M_WDS_IDY As String
        Dim M_WDS_SRL As String
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            If DIR_RS.BOF = False Then
                DIR_RS.MoveFirst
                M_WDS_IDY = Trim(DIR_RS("DIR_ALS"))
            End If
            If Len(M_WDS_IDY) < 8 Then
                MsgBox ("Distributor ID Required for this Process! Obtain Distributor ID and Start Again"), vbCritical, "DAS Version 4.0"
'                cancel
                Exit Sub
            End If
            M_WDS_SRL = (Val(DIR_RS("DIR_MSG")) + 1)
            If Len(Trim(M_WDS_SRL)) < 3 Then
                 While Len(Trim(M_WDS_SRL)) < 3
                      M_WDS_SRL = "0" & M_WDS_SRL
                 Wend
            End If
        End If
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'SYS' and DIR_IDX = 'VER'", con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            If DIR_RS.BOF = False Then
                DIR_RS.MoveFirst
                DIR_RS("DIR_MSG") = Trim("01.01.01")
                DIR_RS.Update
            End If
        End If
' * Distributor ID Identification
        Dim Pending_RS As New ADODB.Recordset
        If Pending_RS.State = 1 Then Pending_RS.Close
        Pending_RS.Open "select * from DIR where DIR_DAT >= CTOD('" & Format(DTPicker4, "DD/MM/YYYY") & "') AND DIR_DAT <= CTOD('" & Format(DTPicker3, "DD/MM/YYYY") & "') AND alltrim(DIR_MSG) = 'Y' AND substr(alltrim(DIR_IDY),1,3) = 'DAY'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
        If Pending_RS.RecordCount > 0 Then
            Pending_RS.MoveFirst
            ST_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
            Pending_RS.MoveLast
            EN_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
        End If

If MsgBox("Create Data File for Data Transfer Upto " & Format(DTPicker3.Value, "DD/MM/YYYY"), vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
        
    ' * : Beginning of PARTY.TXT Creation Code for Data Transfer
    Me.MousePointer = vbHourglass
    Dim F_OMS, F_LOC, F_MAR, F_SMN
    Dim FS_SOU, FS_LOC, FS_MAR, FS_SMN As Object
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    Set FS_LOC = CreateObject("Scripting.filesystemobject")
    Set FS_MAR = CreateObject("SCripting.filesystemobject")
    Set FS_SMN = CreateObject("SCripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\PARTY.TXT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\PARTY.TXT")
    End If
    If FS_LOC.FileExists(CURDIR & "\HHT\LOCATION.TXT") = True Then
        FS_LOC.DeleteFile (CURDIR & "\HHT\LOCATION.TXT")
    End If
    If FS_MAR.FileExists(CURDIR & "\HHT\MARKET.TXT") = True Then
        FS_MAR.DeleteFile (CURDIR & "\HHT\MARKET.TXT")
    End If
    If FS_SMN.FileExists(CURDIR & "\HHT\SALESMAN.TXT") = True Then
        FS_SMN.DeleteFile (CURDIR & "\HHT\SALESMAN.TXT")
    End If
    
    Set F_OMS = FS_SOU.OpenTextFile(CURDIR & "\HHT\PARTY.TXT", 2, True)
    Set F_LOC = FS_LOC.OpenTextFile(CURDIR & "\HHT\LOCATION.TXT", 2, True)
    Set F_MAR = FS_MAR.OpenTextFile(CURDIR & "\HHT\MARKET.TXT", 2, True)
    Set F_SMN = FS_SMN.OpenTextFile(CURDIR & "\HHT\SALESMAN.TXT", 2, True)
    
    Dim J As Integer
    J = 0
    
    
    Dim Category_RS As New ADODB.Recordset
    Dim Location_RS As New ADODB.Recordset
    Dim Market_RS As New ADODB.Recordset
    Dim Salesman_RS As New ADODB.Recordset
    
    Dim Retailer_RS As New ADODB.Recordset
    If Retailer_RS.State = 1 Then Retailer_RS.Close
    Retailer_RS.Open "select * from ACT where ACT_TYP = 'SAL' or ACT_TYP = 'PUR'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh  CST_TYP <> 'Van'
    
    If Retailer_RS.RecordCount > 0 Then Retailer_RS.MoveFirst
       
       Dim L_CAT_NME As String
       Dim L_LOC_NME As String
       Dim L_MAR_NME As String
       Dim L_DST_NME As String
       Dim L_MAR_GRD As String
       
       If Category_RS.State = 1 Then Category_RS.Close
       Category_RS.Open "SELECT * FROM CAT", con, adOpenKeyset, adLockPessimistic
       If Location_RS.State = 1 Then Location_RS.Close
       Location_RS.Open "SELECT * FROM LOC", con, adOpenKeyset, adLockPessimistic
       If Market_RS.State = 1 Then Market_RS.Close
       Market_RS.Open "SELECT * FROM MAR", con, adOpenKeyset, adLockPessimistic
       
       While Not Retailer_RS.EOF
       
          If Category_RS.RecordCount > 0 Then Category_RS.MoveFirst
          Category_RS.Find "CAT_IDY='" & Trim(Retailer_RS("CAT_IDY")) & "'"
          If Category_RS.EOF = False Then
             L_CAT_NME = Category_RS("CAT_NME")
          End If
          
          If Location_RS.RecordCount > 0 Then Location_RS.MoveFirst
          Location_RS.Find "LOC_IDY='" & Trim(Retailer_RS("LOC_IDY")) & "'"
          If Location_RS.EOF = False Then
             L_LOC_NME = Location_RS("LOC_NME")
          End If
          
          If Market_RS.RecordCount > 0 Then Market_RS.MoveFirst
          Market_RS.Find "MAR_IDY='" & Trim(Retailer_RS("MAR_IDY")) & "'"
          If Market_RS.EOF = False Then
             L_MAR_NME = Market_RS("MAR_NME")
             L_DST_NME = Market_RS("MAR_PGR")
             L_MAR_GRD = Market_RS("MAR_GRD")
          End If
          
'          F_OMS.WriteLine Trim(M_WDS_IDY) & "|" & Trim(Retailer_RS("ACT_IDY")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(L_DST_NME) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(L_LOC_NME) & "|" & Trim(L_CAT_NME) & "|" & Trim(L_MAR_NME) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("DBT_TYP"))
          If Trim(Retailer_RS("ACT_TYP")) = "PUR" Then
                F_OMS.WriteLine Trim(Retailer_RS("ACT_IDY")) & "|" & "S" & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(Retailer_RS("LOC_IDY")) & "|" & Trim(Retailer_RS("MAR_IDY")) & "|" & Trim(Retailer_RS("CAT_IDY")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("DBT_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("ROU_IDY"))
                J = J + 1
          Else
                F_OMS.WriteLine Trim(Retailer_RS("ACT_IDY")) & "|" & "C" & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(Retailer_RS("LOC_IDY")) & "|" & Trim(Retailer_RS("MAR_IDY")) & "|" & Trim(Retailer_RS("CAT_IDY")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("DBT_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("ROU_IDY"))
                J = J + 1
          End If
          Retailer_RS.MoveNext
       Wend
'''       MsgBox "" & J & ": Outlet Records Created for Data Transfer!"
'''    Else
'''       MsgBox ("No Outlet Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
'''       cancel
'''    End If
       Dim n As Integer
       n = 0
       If Location_RS.State = 1 Then Location_RS.Close
       Location_RS.Open "SELECT * FROM LOC", con, adOpenKeyset, adLockPessimistic
       While Not Location_RS.EOF
          F_LOC.WriteLine Trim(Location_RS("LOC_IDY")) & "|" & Trim(Location_RS("LOC_NME"))
          Location_RS.MoveNext
          n = n + 1
       Wend
       
       Dim o As Integer
       o = 0
       If Market_RS.State = 1 Then Market_RS.Close
       Market_RS.Open "SELECT * FROM MAR", con, adOpenKeyset, adLockPessimistic
       While Not Market_RS.EOF
          F_MAR.WriteLine Trim(Market_RS("MAR_IDY")) & "|" & Trim(Market_RS("MAR_NME")) & "|" & Trim(Market_RS("MAR_PGR")) & "|" & Trim(Market_RS("MAR_GRD"))
          Market_RS.MoveNext
          o = o + 1
       Wend
       
       Dim P As Integer
       P = 0
       If Salesman_RS.State = 1 Then Salesman_RS.Close
       Salesman_RS.Open "SELECT * FROM SMN", con, adOpenKeyset, adLockPessimistic
       While Not Salesman_RS.EOF()
          F_SMN.WriteLine Trim(Salesman_RS("SMN_IDY")) & "|" & Trim(Salesman_RS("SMN_NME"))
          Salesman_RS.MoveNext
          P = P + 1
       Wend
       
        F_OMS.Close
        F_LOC.Close
        F_MAR.Close
        F_SMN.Close
'    cancel
    
    ' * : End of PARTY.TXT Creation Code for Data Transfer
    
    ' * : Beginning of STK_TRN.TXT Creation Code for Data Transfer
    
    Dim F_INV
    Dim F_INL
    Dim F_VAN
    Dim F_EIS
    
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    Set FS_INL = CreateObject("Scripting.filesystemobject")
    Set FS_VAN = CreateObject("Scripting.filesystemobject")
    Set FS_EIS = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\STK_TRN.TXT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\STK_TRN.TXT")
    End If
    
    If FS_INL.FileExists(CURDIR & "\HHT\STK_TRN_AMT.TXT") = True Then
        FS_INL.DeleteFile (CURDIR & "\HHT\STK_TRN_AMT.TXT")
    End If
    
    If FS_VAN.FileExists(CURDIR & "\HHT\LOAD_SHEET.TXT") = True Then
        FS_VAN.DeleteFile (CURDIR & "\HHT\LOAD_SHEET.TXT")
    End If
    
    If FS_EIS.FileExists(CURDIR & "\HHT\EMPTY_ISSUE.TXT") = True Then
        FS_EIS.DeleteFile (CURDIR & "\HHT\EMPTY_ISSUE.TXT")
    End If
    
    Set F_INV = FS_SOU.OpenTextFile(CURDIR & "\HHT\STK_TRN.TXT", 2, True)
    Set F_INL = FS_INL.OpenTextFile(CURDIR & "\HHT\STK_TRN_AMT.TXT", 2, True)
    Set F_VAN = FS_VAN.OpenTextFile(CURDIR & "\HHT\LOAD_SHEET.TXT", 2, True)
    Set F_EIS = FS_EIS.OpenTextFile(CURDIR & "\HHT\EMPTY_ISSUE.TXT", 2, True)
    
    Dim K As Double
    Dim r As Double
    Dim s As Double
    Dim T As Double
    K = 0
    r = 0
    s = 0
    T = 0
    Dim INL_RS As New ADODB.Recordset
    Dim INV_RS As New ADODB.Recordset
    Dim PRD_RS As New ADODB.Recordset
    Dim BAT_RS As New ADODB.Recordset
    Dim SCH_RS As New ADODB.Recordset
    Dim VAN_RS As New ADODB.Recordset
    Dim SNV_RS As New ADODB.Recordset
    Dim ACT_RS As New ADODB.Recordset
    
    Dim L_CST_IDY As String
    Dim L_ORD_IDY As String
    Dim L_BAT_NME As String
    Dim L_PRD_COD As String
    Dim L_MER_NUM As String
    
    If INL_RS.State = 1 Then INL_RS.Close
    If INV_RS.State = 1 Then INV_RS.Close
    If PRD_RS.State = 1 Then PRD_RS.Close
    If BAT_RS.State = 1 Then BAT_RS.Close
    If SCH_RS.State = 1 Then SCH_RS.Close
       
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * FROM DIR where substr(alltrim(DIR_IDY),1,3) = 'DAY' AND alltrim(DIR_MSG) = 'Y' AND DIR_DAT >= CTOD('" & DTPicker4 & "') AND DIR_DAT <= CTOD('" & DTPicker3 & "')", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveFirst
        ST_DAT = DIR_RS("DIR_DAT")
        BEG_DAT = DIR_RS("DIR_DAT")
        DIR_RS.MoveLast
        EN_DAT = DIR_RS("DIR_DAT")
        END_DAT = DIR_RS("DIR_DAT")
    End If
    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    INV_RS.Open "SELECT * FROM INV WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    BAT_RS.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
    SCH_RS.Open "SELECT * FROM SCH", con, adOpenKeyset, adLockPessimistic
    
    If INV_RS.RecordCount > 0 Then INV_RS.MoveFirst
    
        While Not INV_RS.EOF
            
            If PRD_RS.RecordCount > 0 Then PRD_RS.MoveFirst
            PRD_RS.Find "PRD_IDY='" & Trim(INV_RS("PRD_IDY")) & "'"
            If PRD_RS.EOF = False Then
               L_PRD_COD = PRD_RS("ALT_IDY")
            End If
            
            If BAT_RS.RecordCount > 0 Then BAT_RS.MoveFirst
            BAT_RS.Find "BAT_IDY='" & Trim(INV_RS("BAT_IDY")) & "'"
            If BAT_RS.EOF = False Then
               L_BAT_NME = BAT_RS("BAT_NME")
            End If
            
            If INL_RS.RecordCount > 0 Then INL_RS.MoveFirst
            INL_RS.Find "INV_IDY='" & Trim(INV_RS("INV_IDY")) & "'"
            If INL_RS.EOF = False Then
               L_CST_IDY = INL_RS("CST_IDY")
               L_ORD_IDY = INL_RS("ORD_IDY")
               L_LRN_IDY = INL_RS("LRN_IDY")
               L_LRN_DAT = INL_RS("LRN_DAT")
            End If
            
            If SCH_RS.RecordCount > 0 Then SCH_RS.MoveFirst
            SCH_RS.Find "SCH_IDY='" & Trim(INV_RS("SCH_IDY")) & "'"
            If SCH_RS.EOF = False Then
                L_MER_NUM = SCH_RS("MER_NUM")
            End If
            
            'M_WDS_IDY & "|"Trim(STR(Inv_RS("COL_AMT"))) & "|" &
            If Trim(INV_RS("TXN_TYP")) <> "OPS" Then
                F_INV.WriteLine Trim(INV_RS("INV_IDY")) & "|" & Format(INV_RS("INV_DAT"), "DD/MM/YYYY") & "|" & Trim(INV_RS("TXN_TYP")) & "|" & Trim(L_CST_IDY) & "|" & Trim(L_PRD_COD) & "|" & QTY(INV_RS("PRD_QTY")) & "|" & amt(INV_RS("PRD_RAT")) & "|" & Trim(L_BAT_NME) & "|" & Format(BAT_RS("BRD_DOM"), "DD/MM/YYYY") & "|" & amt(BAT_RS("PRD_MRP")) & "|" & Trim(INV_RS("STK_TYP")) & "|" & Trim(L_MER_NUM) & "|" & Trim(L_LRN_IDY) & "|" & Format(L_LRN_DAT, "DD/MM/YYYY") & "|" & Trim(INL_RS("TXN_TPT"))
                K = K + 1
            End If
            INV_RS.MoveNext
        Wend
'        MsgBox "" & K & ": Stock Records Created for Data Transfer!"
'    Else
'       MsgBox ("No Stock Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
'       cancel
'    End If
    F_INV.Close
    
    If INL_RS.State = 1 Then INL_RS.Close
    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    If INL_RS.RecordCount > 0 Then INL_RS.MoveFirst
    While Not INL_RS.EOF
        M_ADD_AMT = 0
        M_DIS_AMT = 0
        M_TAX_AMT = 0
        If INV_RS.State = 1 Then INV_RS.Close
        INV_RS.Open "SELECT * FROM INV WHERE INV.INV_IDY = '" & Trim(INL_RS("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
        If INV_RS.RecordCount > 0 Then INV_RS.MoveFirst
        While Not INV_RS.EOF
            If IsNull(INV_RS("ADD_AMT")) = False Then
                M_ADD_AMT = M_ADD_AMT + INV_RS("ADD_AMT")
            End If
            If IsNull(INV_RS("DIS_AMT")) = False Then
                M_DIS_AMT = M_DIS_AMT + INV_RS("DIS_AMT")
            End If
            If IsNull(INV_RS("TAX_AMT")) = False Then
                M_TAX_AMT = M_TAX_AMT + INV_RS("TAX_AMT")
            End If
            INV_RS.MoveNext
        Wend
        If Trim(INL_RS("TXN_TYP")) <> "OPS" Then
            If INV_RS.RecordCount > 0 Then
                If Trim(INL_RS("TXN_TYP")) = "PUR" Then
                    F_INL.WriteLine Trim(INL_RS("INV_IDY")) & "|" & Format(INL_RS("INV_DAT"), "DD/MM/YYYY") & "|" & INL_RS("TXN_TYP") & "|" & INL_RS("CST_IDY") & "|" & amt(INL_RS("TAX_AMT") + INL_RS("EXD_AMT") + M_TAX_AMT) & "|" & amt(INL_RS("SPD_AMT") + INL_RS("CSD_AMT") + M_ADD_AMT + M_DIS_AMT) & "|" & amt(INL_RS("NET_AMT")) & "|" & Trim(INL_RS("SMN_IDY")) & "|" & Trim(INL_RS("ROU_IDY")) & "|" & Trim(INL_RS("ORD_IDY"))
                    r = r + 1
                Else
                    F_INL.WriteLine Trim(INL_RS("INV_IDY")) & "|" & Format(INL_RS("INV_DAT"), "DD/MM/YYYY") & "|" & INL_RS("TXN_TYP") & "|" & INL_RS("CST_IDY") & "|" & amt(INL_RS("TAX_AMT") + INL_RS("EXD_AMT") + M_TAX_AMT) & "|" & amt(INL_RS("SPD_AMT") + INL_RS("CSD_AMT") + M_ADD_AMT + M_DIS_AMT) & "|" & amt(INL_RS("NET_AMT")) & "|" & Trim(INL_RS("SMN_IDY")) & "|" & Trim(INL_RS("ROU_IDY")) & "|" & ""
                    r = r + 1
                End If
            End If
        End If
        INL_RS.MoveNext
    Wend
    F_INL.Close
    
'    cancel

    If VAN_RS.State = 1 Then VAN_RS.Close
    VAN_RS.Open "SELECT * FROM VAN WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    If VAN_RS.RecordCount > 0 Then VAN_RS.MoveFirst
    While Not VAN_RS.EOF
        If Trim(VAN_RS("VAN_CLS")) = "C" Then
            If PRD_RS.State = 1 Then PRD_RS.Close
            PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
            If PRD_RS.RecordCount > 0 Then PRD_RS.MoveFirst
                PRD_RS.Find "PRD_IDY='" & Trim(VAN_RS("PRD_IDY")) & "'"
                If PRD_RS.EOF = False Then
                    L_PRD_COD = PRD_RS("ALT_IDY")
                End If
        
            F_VAN.WriteLine Format(VAN_RS("INV_DAT"), "DD/MM/YYYY") & "|" & Trim(VAN_RS("ROU_IDY")) & "|" & Trim(VAN_RS("SMN_IDY")) & "|" & Trim(VAN_RS("INV_IDY")) & "|" & Trim(L_PRD_COD) & "|" & QTY(VAN_RS("TOT_QTY")) & "|" & QTY(VAN_RS("RET_QTY")) & "|" & QTY(VAN_RS("BRK_QTY")) & "|" & QTY(VAN_RS("LKG_QTY")) & "|" & VAN_RS("CST_IDY")
            s = s + 1
        End If
        VAN_RS.MoveNext
    Wend
    F_VAN.Close
    
    If SNV_RS.State = 1 Then SNV_RS.Close
    SNV_RS.Open "SELECT * FROM SNV WHERE TXN_DAT >= CTOD('" & ST_DAT & "') .AND. TXN_DAT <= CTOD('" & EN_DAT & "') AND (TXN_TYP = 'PUR' OR TXN_TYP = 'AOU' OR TXN_TYP = 'PRT' OR TXN_TYP = 'STI' OR TXN_TYP = 'STO')", con, adOpenKeyset, adLockPessimistic
    If SNV_RS.RecordCount > 0 Then SNV_RS.MoveFirst
    While Not SNV_RS.EOF
        If ACT_RS.State = 1 Then ACT_RS.Close
            ACT_RS.Open "SELECT * FROM ACT where ACT_TYP = 'PUR'", con, adOpenKeyset, adLockPessimistic
            If ACT_RS.RecordCount > 0 Then ACT_RS.MoveFirst
                ACT_RS.Find "ACT_IDY='" & Trim(SNV_RS("CST_IDY")) & "'"
                If ACT_RS.EOF = False Then
                    If Len(Trim(SNV_RS("TXN_IDY"))) = 0 Then
                        If SNV_RS("R_SHL_NOS") + SNV_RS("D_SHL_NOS") > 0 Or SNV_RS("R_PCK_NOS") + SNV_RS("D_PCK_NOS") > 0 Then
                            F_EIS.WriteLine Trim(SNV_RS("INV_IDY")) & "|" & Format(SNV_RS("TXN_DAT"), "DD/MM/YYYY") & "|" & Trim(SNV_RS("PRD_PCK")) & "|" & SNV_RS("R_SHL_NOS") + SNV_RS("D_SHL_NOS") & "|" & QTY(STM(SNV_RS("R_PCK_NOS") + SNV_RS("D_PCK_NOS"), SNV_RS("PRD_UPC"))) & "|" & Trim(SNV_RS("TXN_TYP"))
                            T = T + 1
                        End If
                    Else
                        If SNV_RS("R_SHL_NOS") + SNV_RS("D_SHL_NOS") > 0 Or SNV_RS("R_PCK_NOS") + SNV_RS("D_PCK_NOS") > 0 Then
                            F_EIS.WriteLine Trim(SNV_RS("TXN_IDY")) & "|" & Format(SNV_RS("TXN_DAT"), "DD/MM/YYYY") & "|" & Trim(SNV_RS("PRD_PCK")) & "|" & SNV_RS("R_SHL_NOS") + SNV_RS("D_SHL_NOS") & "|" & QTY(STM(SNV_RS("R_PCK_NOS") + SNV_RS("D_PCK_NOS"), SNV_RS("PRD_UPC"))) & "|" & Trim(SNV_RS("TXN_TYP"))
                            T = T + 1
                        End If
                    End If
                End If
        SNV_RS.MoveNext
    Wend
    
    ' * : End of STK_TRN.TXT, LOAD_SHEET.TXT, EMPTY_ISSUE.TXT Creation Code for Data Transfer
    ' * : Beginning of STOCK.TXT Creation Code for Data Transfer

    Dim F_STK
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\STOCK.TXT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\STOCK.TXT")
    End If
    
    Set F_STK = FS_SOU.OpenTextFile(CURDIR & "\HHT\STOCK.TXT", 2, True)
    Dim l As Integer
    l = 0
    LDD = DateDiff("D", BEG_DAT, END_DAT)
    LDD = LDD + 1
    Do While LDD <> 0
            STKTXTDEX (BEG_DAT)
            Dim CLD_RS As New ADODB.Recordset
            
            If CLD_RS.State = 1 Then CLD_RS.Close
            If PRD_RS.State = 1 Then PRD_RS.Close
            If BAT_RS.State = 1 Then BAT_RS.Close
            
            CLD_RS.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockPessimistic
            PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
            BAT_RS.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
            
            
            If CLD_RS.RecordCount > 0 Then CLD_RS.MoveFirst
            
                While Not CLD_RS.EOF
                               
                    If BAT_RS.RecordCount > 0 Then BAT_RS.MoveFirst
                    BAT_RS.Find "BAT_IDY='" & Trim(CLD_RS("PRD_IDY")) & "'"
                    If BAT_RS.EOF = False Then
                       L_BAT_NME = BAT_RS("BAT_NME")
                    End If
                    
                    If PRD_RS.RecordCount > 0 Then PRD_RS.MoveFirst
                    PRD_RS.Find "PRD_IDY='" & Trim(BAT_RS("PRD_IDY")) & "'"
                    If PRD_RS.EOF = False Then
                       L_PRD_COD = PRD_RS("ALT_IDY")
                    End If
                                
        '            F_STK.WriteLine Trim(M_WDS_IDY) & "|" & Trim(L_PRD_COD) & "|" & Trim(STR(CLD_RS("OPN_STK"))) & "|" & Trim(STR(CLD_RS("PUR_STK"))) & "|" & Trim(STR(CLD_RS("PRT_STK"))) & "|" & Trim(STR(CLD_RS("SAL_STK"))) & "|" & Trim(STR(CLD_RS("SRN_STK"))) & "|" & Trim(STR(CLD_RS("STI_STK"))) & "|" & Trim(STR(CLD_RS("STO_STK"))) & "|" & Trim(STR(CLD_RS("CLS_STK"))) & "|" & Trim(STR(BAT_RS("PRD_MRP"))) & "|" & Trim(BAT_RS("BAT_NME")) & "|" & BAT_RS("BRD_DOM")
                    F_STK.WriteLine Format(BEG_DAT, "DD/MM/YYYY") & "|" & Trim(L_PRD_COD) & "|" & Trim(L_BAT_NME) & "|" & BAT_RS("BRD_DOM") & "|" & amt(BAT_RS("PRD_MRP")) & "|" & QTY(STM(CLD_RS("OPN_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("PUR_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("PRT_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("SAL_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("SRN_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("STI_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("STO_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("CLS_STK"), PRD_RS("SUB_UNT")))
                    CLD_RS.MoveNext
                    l = l + 1
                Wend
        '        MsgBox "" & L & ": SKU Records Created for Data Transfer!"
        '    Else
        '       MsgBox ("No Stock Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
        '       cancel
        '    End If
           
        '    cancel
            ' * : End of STOCK.TXT Creation Code for Data Transfer
        BEG_DAT = (BEG_DAT + 1)
        LDD = LDD - 1
    Loop
    F_STK.Close
    
    ' * : Beginning of EMPTIES.TXT Creation Code for Data Transfer
    ' * : First Clear NULL values from SNV Table
    'Dim SNV_RS As New ADODB.Recordset
    'Dim CLD_RS As New ADODB.Recordset
    If SNV_RS.State = 1 Then SNV_RS.Close
'    SNV_RS.Open "SELECT TXN_TYP, PRD_IDY, I_SHL_NOS, I_PCK_NOS, DP_SHL_NOS, DP_PCK_NOS, DEP_AMT, R_SHL_NOS, R_PCK_NOS, BRK_SHL, BRK_PCK, D_SHL_NOS, D_PCK_NOS FROM SNV", CON, adOpenKeyset, adlockpessimistic
    SNV_RS.Open "SELECT * FROM SNV", con, adOpenKeyset, adLockPessimistic
'    If SNV_RS.RecordCount > 0 Then SNV_RS.MoveFirst
    While SNV_RS.EOF = False
            If IsNull(SNV_RS("I_SHL_NOS")) = True Then
                    SNV_RS("I_SHL_NOS") = "0"
                    SNV_RS.Update
            End If
            If IsNull(SNV_RS("I_PCK_NOS")) = True Then
                SNV_RS("I_PCK_NOS") = "0"
                SNV_RS.Update
            End If
            If IsNull(SNV_RS("DP_SHL_NOS")) = True Then
                SNV_RS("DP_SHL_NOS") = "0"
                SNV_RS.Update
            End If
            If IsNull(SNV_RS("DP_PCK_NOS")) = True Then
                SNV_RS("DP_PCK_NOS") = "0"
                SNV_RS.Update
            End If
            If IsNull(SNV_RS("DEP_AMT")) = True Then
                SNV_RS("DEP_AMT") = "0"
                SNV_RS.Update
            End If
            If IsNull(SNV_RS("R_SHL_NOS")) = True Then
                SNV_RS("R_SHL_NOS") = "0"
                SNV_RS.Update
            End If
            If IsNull(SNV_RS("R_PCK_NOS")) = True Then
                SNV_RS("R_PCK_NOS") = "0"
                SNV_RS.Update
            End If
            If IsNull(SNV_RS("BRK_SHL")) = True Then
                SNV_RS("BRK_SHL") = "0"
                SNV_RS.Update
            End If
            If IsNull(SNV_RS("BRK_PCK")) = True Then
                SNV_RS("BRK_PCK") = "0"
                SNV_RS.Update
            End If
            If IsNull(SNV_RS("D_SHL_NOS")) = True Then
                SNV_RS("D_SHL_NOS") = "0"
                SNV_RS.Update
            End If
            If IsNull(SNV_RS("D_PCK_NOS")) = True Then
                SNV_RS("D_PCK_NOS") = "0"
                SNV_RS.Update
            End If
            SNV_RS.MoveNext
    Wend
    ' * : Null Values Cleared and Updated into SNV
    
    Dim F_EMT
    Dim F_ACG
    Dim F_PAY
    
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    Set FS_ACG = CreateObject("Scripting.filesystemobject")
    Set FS_PAY = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\EMPTIES.TXT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\EMPTIES.TXT")
    End If
    
    If FS_ACG.FileExists(CURDIR & "\HHT\ARCOLGOD.TXT") = True Then
        FS_ACG.DeleteFile (CURDIR & "\HHT\ARCOLGOD.TXT")
    End If
    
    If FS_PAY.FileExists(CURDIR & "\HHT\ACT_PAY.TXT") = True Then
        FS_PAY.DeleteFile (CURDIR & "\HHT\ACT_PAY.TXT")
    End If
    
    Set F_EMT = FS_SOU.OpenTextFile(CURDIR & "\HHT\EMPTIES.TXT", 2, True)
    Set F_ACG = FS_ACG.OpenTextFile(CURDIR & "\HHT\ARCOLGOD.TXT", 2, True)
    Set F_PAY = FS_PAY.OpenTextFile(CURDIR & "\HHT\ACT_PAY.TXT", 2, True)
    
    Dim M As Integer
    Dim V As Integer
    M = 0
    V = 0
    BEG_DAT = ST_DAT
    END_DAT = EN_DAT
    LDD = DateDiff("D", BEG_DAT, END_DAT)
    LDD = LDD + 1
    Do While LDD <> 0
        EMPTXTDEX (BEG_DAT)
        If CLD_RS.State = 1 Then CLD_RS.Close
        CLD_RS.Open "SELECT * FROM RP1", con, adOpenKeyset, adLockPessimistic
        If CLD_RS.RecordCount > 0 Then CLD_RS.MoveFirst
        While Not CLD_RS.EOF
            F_EMT.WriteLine Format(BEG_DAT, "DD/MM/YYYY") & "|" & Trim(CLD_RS("CHR_TEN")) & "|" & CLD_RS("NUM_ONE") & "|" & QTY(STM((CLD_RS("NUM_TWO")), CLD_RS("NUM_THR")))
            M = M + 1
            CLD_RS.MoveNext
        Wend
        BEG_DAT = (BEG_DAT + 1)
        LDD = LDD - 1
    Loop
    F_EMT.Close
'    MsgBox "" & M & ": Glass Records Created for Data Transfer!"
'    End If
'    cancel
' - * ARCOLGOD.TXT
    BEG_DAT = ST_DAT
    END_DAT = EN_DAT
    LDD = DateDiff("D", BEG_DAT, END_DAT)
    LDD = LDD + 1
    Do While LDD <> 0
        COLGODDEX (BEG_DAT)
        If CLD_RS.State = 1 Then CLD_RS.Close
        CLD_RS.Open "SELECT * FROM RP1", con, adOpenKeyset, adLockPessimistic
        If CLD_RS.RecordCount > 0 Then CLD_RS.MoveFirst
        While Not CLD_RS.EOF
            F_ACG.WriteLine Format(BEG_DAT, "DD/MM/YYYY") & "|" & Trim(CLD_RS("CHR_TEN")) & "|" & CLD_RS("NUM_ONE") & "|" & QTY(STM((CLD_RS("NUM_TWO")), CLD_RS("NUM_THR"))) & "|" & CLD_RS("NUM_FOR") & "|" & QTY(STM((CLD_RS("NUM_FIV")), CLD_RS("NUM_THR")))
            V = V + 1
            CLD_RS.MoveNext
        Wend
        BEG_DAT = (BEG_DAT + 1)
        LDD = LDD - 1
    Loop
    F_ACG.Close
'    MsgBox "" & M & ": Glass Records Created for Data Transfer!"
'    End If
'    cancel
' * : End of EMPTIES.TXT, ARCOLGOD.TXT Creation Code for Data Transfer
    
    Dim W As Integer
    W = 0
    BEG_DAT = ST_DAT
    END_DAT = EN_DAT
    LDD = DateDiff("D", BEG_DAT, END_DAT)
    LDD = LDD + 1
    Do While LDD <> 0
        ACTPAYDEX (BEG_DAT)
        If CLD_RS.State = 1 Then CLD_RS.Close
        CLD_RS.Open "SELECT * FROM OPA", con, adOpenKeyset, adLockPessimistic
        If CLD_RS.RecordCount > 0 Then CLD_RS.MoveFirst
        While Not CLD_RS.EOF
            If ACT_RS.State = 1 Then ACT_RS.Close
            ACT_RS.Open "SELECT * FROM ACT WHERE ALLTRIM(ACT_IDY) = '" & Trim(CLD_RS("ACT_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
            If ACT_RS.RecordCount > 0 Then ACT_RS.MoveFirst
            If Trim(ACT_RS("ACT_TYP")) = "PUR" Then
               F_PAY.WriteLine Format(BEG_DAT, "DD/MM/YYYY") & "|" & Trim(CLD_RS("ACT_IDY")) & "|" & amt(CLD_RS("OPA_DBT") - CLD_RS("OPA_CRT"))
               W = W + 1
            End If
            CLD_RS.MoveNext
        Wend
        BEG_DAT = (BEG_DAT + 1)
        LDD = LDD - 1
    Loop
    F_PAY.Close

    ' * : Log Text File Creation Begin Process
    
    Dim F_VER
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    If FS_SOU.FileExists(CURDIR & "\HHT\VERSION.TXT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\VERSION.TXT")
    End If
    
    Dim Q As Integer
    Dim LDC_DAT As Date
    Q = 0
    Set F_VER = FS_SOU.OpenTextFile(CURDIR & "\HHT\VERSION.TXT", 2, True)
    
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * FROM DIR where substr(alltrim(DIR_IDY),1,3) = 'DAY' AND alltrim(DIR_MSG) = 'Y'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveLast
        LDC_DAT = DIR_RS("DIR_DAT")
    Else
        LDC_DAT = Date
    End If
    F_VER.WriteLine Trim(M_WDS_IDY) & "|" & Trim("01.01.01") & "|" & Format(DTPicker4, "dd/MM/yyyy") & "|" & Format(DTPicker3, "dd/MM/yyyy") & "|" & Format(LDC_DAT, "dd/MM/yyyy")
    Q = Q + 1
    
    F_VER.Close
    
    Dim F_LOG
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
       FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    
    If FS_SOU.FileExists(CURDIR & "\HHT\LOG.TXT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\LOG.TXT")
    End If
    
    Set F_LOG = FS_SOU.OpenTextFile(CURDIR & "\HHT\LOG.TXT", 2, True)
    
    F_LOG.WriteLine "PARTY.TXT" & "|" & str(J)
    F_LOG.WriteLine "LOCATION.TXT" & "|" & str(n)
    F_LOG.WriteLine "MARKET.TXT" & "|" & str(o)
    F_LOG.WriteLine "SALESMAN.TXT" & "|" & str(P)
    F_LOG.WriteLine "STK_TRN.TXT" & "|" & str(K)
    F_LOG.WriteLine "STK_TRN_AMT.TXT" & "|" & str(r)
    F_LOG.WriteLine "LOAD_SHEET.TXT" & "|" & str(s)
    F_LOG.WriteLine "EMPTY_ISSUE.TXT" & "|" & str(T)
    F_LOG.WriteLine "STOCK.TXT" & "|" & str(l)
    F_LOG.WriteLine "EMPTIES.TXT" & "|" & str(M)
    F_LOG.WriteLine "ARCOLGOD.TXT" & "|" & str(V)
    F_LOG.WriteLine "ACT_PAY.TXT" & "|" & str(W)
    F_LOG.WriteLine "VERSION.TXT" & "|" & str(Q)
    
    Dim MDD As String
    Dim MMM As String
    Dim MYY As String
    MDD = Trim(str(Day(DTPicker3)))
    MMM = Trim(str(Month(DTPicker3)))
    MYY = Trim(str(Year(DTPicker3)))
'    F_LOG.WriteLine MDD & MMM & MYY & M_WDS_SRL 'Refers to Day Closure Date, Last Line in Log File is Blocked - PNR
    F_LOG.Close
    ' * : Log Text File Creation End Process
    Me.MousePointer = vbDefault
    MsgBox ("File Creation Process Successful! Log File Generated!"), vbInformation, "DAS Version 4.0"
    
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then DIR_RS.MoveFirst
       DIR_RS("DIR_MSG") = Trim(M_WDS_SRL)
       DIR_RS.Update
    If DIR_RS.State = 1 Then DIR_RS.Close
    'Zip File Creation
            Dim DEX As String
            DEX = "C:\DEX"
            If Dir(DEX, vbDirectory) = "" Then
                MkDir DEX
            End If
            
            CopyFile CURDIR & "\HHT" & "\PARTY.TXT", DEX & "\PARTY.TXT", 0
            CopyFile CURDIR & "\HHT" & "\MARKET.TXT", DEX & "\MARKET.TXT", 0
            CopyFile CURDIR & "\HHT" & "\LOCATION.TXT", DEX & "\LOCATION.TXT", 0
            CopyFile CURDIR & "\HHT" & "\SALESMAN.TXT", DEX & "\SALESMAN.TXT", 0
            CopyFile CURDIR & "\HHT" & "\STK_TRN.TXT", DEX & "\STK_TRN.TXT", 0
            CopyFile CURDIR & "\HHT" & "\STK_TRN_AMT.TXT", DEX & "\STK_TRN_AMT.TXT", 0
            CopyFile CURDIR & "\HHT" & "\LOAD_SHEET.TXT", DEX & "\LOAD_SHEET.TXT", 0
            CopyFile CURDIR & "\HHT" & "\EMPTY_ISSUE.TXT", DEX & "\EMPTY_ISSUE.TXT", 0
            CopyFile CURDIR & "\HHT" & "\STOCK.TXT", DEX & "\STOCK.TXT", 0
            CopyFile CURDIR & "\HHT" & "\EMPTIES.TXT", DEX & "\EMPTIES.TXT", 0
            CopyFile CURDIR & "\HHT" & "\ARCOLGOD.TXT", DEX & "\ARCOLGOD.TXT", 0
            CopyFile CURDIR & "\HHT" & "\ACT_PAY.TXT", DEX & "\ACT_PAY.TXT", 0
            CopyFile CURDIR & "\HHT" & "\VERSION.TXT", DEX & "\VERSION.TXT", 0
            CopyFile CURDIR & "\HHT" & "\LOG.TXT", DEX & "\LOG.TXT", 0
            CopyFile CURDIR & "\HHT" & "\LHA.EXE", DEX & "\LHA.EXE", 0
            CopyFile CURDIR & "\HHT" & "\FCR.EXE", DEX & "\FCR.EXE", 0
            
            CopyFile App.Path & "\" & "Zip32.dll", CURDIR & "\" & "Zip32.dll", 0
            
            M_FIL_NME = "\" & "DAS" & Trim(M_WDS_IDY) & Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & ".ZIP"
            Dim Myzip As ColaZip.GSColaZip32
            Set Myzip = New ColaZip.GSColaZip32
            Myzip.RootDirectory = "\"
            'myzip.ZipFileName = App.Path & "\text1.zip"
            'myzip.AddFile (App.Path & "\*.*")
            Myzip.ZipFileName = "C:\DEX" & M_FIL_NME
'            Myzip.AddFile ("C:\DEX" & "\*.TXT")

            Myzip.AddFile ("C:\DEX" & "\PARTY.TXT")
            Myzip.AddFile ("C:\DEX" & "\MARKET.TXT")
            Myzip.AddFile ("C:\DEX" & "\LOCATION.TXT")
            Myzip.AddFile ("C:\DEX" & "\SALESMAN.TXT")
            Myzip.AddFile ("C:\DEX" & "\STK_TRN.TXT")
            Myzip.AddFile ("C:\DEX" & "\STK_TRN_AMT.TXT")
            Myzip.AddFile ("C:\DEX" & "\LOAD_SHEET.TXT")
            Myzip.AddFile ("C:\DEX" & "\EMPTY_ISSUE.TXT")
            Myzip.AddFile ("C:\DEX" & "\STOCK.TXT")
            Myzip.AddFile ("C:\DEX" & "\EMPTIES.TXT")
            Myzip.AddFile ("C:\DEX" & "\ARCOLGOD.TXT")
            Myzip.AddFile ("C:\DEX" & "\ACT_PAY.TXT")
            Myzip.AddFile ("C:\DEX" & "\VERSION.TXT")
            Myzip.AddFile ("C:\DEX" & "\LOG.TXT")

            Myzip.MakeColaZip ("DAS")
            
            'Shell DEX & "\FCR.EXE", vbHide
            MsgBox ("Zip File " & M_FIL_NME & " Created in the Folder " & DEX), vbInformation, "DAS Version 4.0"
            M_MAIL_ATTMNT = "C:\DEX" & M_FIL_NME
            If MsgBox("Send Zip File Data to DAS Desk Now ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
'                MailSend "info@consoftsystems.com", "DAS Data Xfer", "DAS Data Zip!", Trim(M_MAIL_ATTMNT)
                MailSend "effmcg@consoftsystems.com", "DAS Data Xfer", "DAS Data Zip from DAS Go Live Distributor!", Trim(M_MAIL_ATTMNT)
            Else
                cancel
            End If
    ' Zip File Creation
            
    cancel
Else
    DTPicker4.SetFocus
End If
End Sub

Private Sub Command3_Click()
If DTPicker4 > DTPicker3 Then
    MsgBox "'From' Date Should Not Be Greater Than 'To' Date", vbExclamation, "DAS Version 4.0"
    DTPicker4.SetFocus
    Exit Sub
End If
If LVMOD2.ListItems.count = 0 Then
    MsgBox "Selection of Days for File Mail is Required!", vbInformation, "DAS Version 4.0"
    DTPicker4.SetFocus
Else
    'DEXFILE
    MM_EX_FILE      'DAS Margin Minder Extract by PNR
End If
End Sub

Private Sub Command4_Click()
If MsgBox("Update all RGB SKUs with this Crate Rental, VAT% Values ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
   CRT_MEX                     ' Update Existing SKUs in PRD.DBF, fields, LST_RAT, LST_DIS
Else
   cancel
End If
End Sub

Private Sub DTPicker1_CloseUp()
DTPicker1.SetFocus
End Sub

Private Sub DTPicker1_GotFocus()
DTPicker1.Value = Date
SETYEAR M_FRM_EXG
DTPicker1.SetFocus
LVMOD.ListItems.CLEAR
LVMOD2.ListItems.CLEAR
Command1.Caption = "&Select All"
End Sub
Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command17.SetFocus
End If
End Sub

Private Sub DTPicker3_CloseUp()
DTPicker3.SetFocus
End Sub

Private Sub DTPicker3_GotFocus()
LVMOD2.ListItems.CLEAR
End Sub

Private Sub DTPicker3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If LVMOD2.ListItems.count > 0 Then  'Previous Buffer Clear
        LVMOD2.ListItems.CLEAR
    End If
    Dim Pending_RS As New ADODB.Recordset
    If Pending_RS.State = 1 Then Pending_RS.Close
    'Pending_RS.Open "select * from DIR where DIR_DAT >= CTOD('" & Format(DTPicker4, "DD/MM/YYYY") & "') AND DIR_DAT <= CTOD('" & Format(DTPicker3, "DD/MM/YYYY") & "') AND alltrim(DIR_MSG) = 'Y' AND substr(alltrim(DIR_IDY),1,3) = 'DAY'", CON, adOpenKeyset, adlockpessimistic  'Modified by Ramanesh
    Pending_RS.Open "select * from DIR where alltrim(DIR_MSG) = 'Y' AND substr(alltrim(DIR_IDY),1,3) = 'DAY'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
    If Pending_RS.RecordCount > 0 Then
    Pending_RS.MoveFirst
        For i = 1 To Pending_RS.RecordCount
            If CDate(Trim(Pending_RS("DIR_DAT"))) >= DTPicker4.Value And CDate(Trim(Pending_RS("DIR_DAT"))) <= DTPicker3.Value Then
                LVMOD2.ListItems.Add = Format(Trim(Pending_RS("DIR_DAT")), "DD/MM/YYYY")
            End If
            Pending_RS.MoveNext
        Next
        LVMOD2.SetFocus
        If LVMOD2.ListItems.count > 0 Then
            LVMOD2.ListItems(1).Selected = True
            LVMOD2.SetFocus
        Else
            MsgBox ("There are no Day Closures Between These Dates"), vbInformation, "DAS Version 4.0"
            cancel
        End If
        Else
            MsgBox ("There are no Day Closures Between These Date"), vbInformation, "DAS Version 4.0"
            cancel
    End If
ElseIf KeyCode <> 27 Then
    DTPicker3.SetFocus
End If
End Sub
Private Sub DTPicker3_LostFocus()
If DTPicker4 > DTPicker3 Then
    MsgBox "'From' Date Should Not Be Greater Than 'To' Date", vbExclamation, "DAS Version 4.0"
    DTPicker4.SetFocus
    Exit Sub
End If
If Month(DTPicker4.Value) <> Month(DTPicker3.Value) Then
    MsgBox "Date Selection Should be From a Single Month", vbCritical, "DAS Version 4.0"
    DTPicker4.SetFocus
    Exit Sub
End If
End Sub

Private Sub DTPicker4_CloseUp()
DTPicker4.SetFocus
End Sub

Private Sub DTPicker4_GotFocus()
LVMOD2.ListItems.CLEAR
End Sub

Private Sub DTPicker4_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
DTPicker3.SetFocus
ElseIf K <> 27 Then
DTPicker4.SetFocus
End If

End Sub

Private Sub DTPicker4_LostFocus()
If DTPicker4 > DTPicker3 Then
    MsgBox "'From' Date Should Not Be Greater Than 'To' Date", vbExclamation, "DAS Version 4.0"
    DTPicker4.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
STP_VAL
' Data Import-Export Folder Creation
Dim DEX As String
Dim CDI As String
Dim CDO As String
DEX = "C:\DEX"
CDI = "C:\DEX\INBOX"
CDO = "C:\DEX\OUTBOX"
If Dir(DEX, vbDirectory) = "" Then
   MkDir DEX
End If
If Dir(CDI, vbDirectory) = "" Then
   MkDir CDI
End If
If Dir(CDO, vbDirectory) = "" Then
   MkDir CDO
End If
' Data Import-Export Folder Creation

SSTab1.Tab = 0
Command2.Enabled = False
DTPicker1.Value = Date
DTPicker1 = Format(DTPicker1, "DD/MM/YYYY")
SETYEAR M_FRM_EXG
DTPicker1 = Format(DTPicker1, "mm/dd/yyyy")
DTPicker1.SetFocus
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
Unload Me
End If
End Sub

Private Sub cancel()
LVMOD.ListItems.CLEAR
LVMOD2.ListItems.CLEAR
Command1.Caption = "&Select All"
DTPicker1.SetFocus
End Sub
Private Sub M_CRT_RNT_Change()
If Len(Trim(M_CRT_RNT)) = 9 Then SendKeys "{home}+{end}"
End Sub
Private Sub M_CRT_RNT_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter Crate - Shell Rental Amount"
End Sub
Private Sub M_CRT_RNT_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    M_CRT_RNT = convert(amt(M_CRT_RNT))
    M_CRT_VAT.SetFocus
ElseIf K = 38 Then
    M_CRT_RNT = convert(amt(M_CRT_RNT))
    M_CRT_VAT.SetFocus
End If
End Sub
Private Sub M_CRT_RNT_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub M_CRT_RNT_LostFocus()
M_CRT_VAT.SetFocus
End Sub

Private Sub M_CRT_VAT_Change()
If Len(Trim(M_CRT_VAT)) = 5 Then SendKeys "{home}+{end}"
End Sub

Private Sub M_CRT_VAT_GotFocus()
SendKeys "{home}+{end}"
M_MDI.STS_BAR.Panels(1).Text = "Enter VAT% on Crate - Shell Rental Amount"
End Sub

Private Sub M_CRT_VAT_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    M_CRT_VAT = convert(amt(M_CRT_VAT))
    Command4.SetFocus
ElseIf K = 38 Then
    M_CRT_VAT = convert(amt(M_CRT_VAT))
    Command4.SetFocus
End If

End Sub

Private Sub M_CRT_VAT_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub M_CRT_VAT_LostFocus()
If M_CRT_VAT >= 100 Then
    MsgBox ("Invalid VAT % Value, Re-Enter"), vbCritical, "DAS Version 4.0"
    M_CRT_VAT.SetFocus
End If
End Sub

Private Sub SSTab1_GotFocus()
If SSTab1.Tab = 1 Then
    Check1.SetFocus
End If
End Sub

Private Sub Text1_Change()
If Len(Trim(M_SPD_AMT)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "e" Then
Unload Me
End If
End Sub
Private Sub EMPTXTDEX(DEN_DATE As Date)
EN_DAT = DEN_DATE
Dim SNV_RS As New ADODB.Recordset
Dim CLD_RS As New ADODB.Recordset
    If SNV_RS.State = 1 Then SNV_RS.Close
    If CLD_RS.State = 1 Then CLD_RS.Close
    con.Execute "DELETE FROM RP1"
    CLD_RS.Open "SELECT * FROM RP1", con, adOpenKeyset, adLockPessimistic
    SNV_RS.Open "SELECT * FROM SNV WHERE TXN_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    If SNV_RS.RecordCount > 0 Then SNV_RS.MoveFirst
    While Not SNV_RS.EOF
        If CLD_RS.State = 1 Then CLD_RS.Close
        CLD_RS.Open "SELECT * FROM RP1 WHERE CHR_TEN = '" & Trim(SNV_RS("PRD_PCK")) & "'", con, adOpenKeyset, adLockPessimistic
        CLD_RS.Find "CHR_TEN='" & Trim(SNV_RS("PRD_PCK")) & "'"
        If CLD_RS.EOF = False Then
            If SNV_RS("TXN_TYP") = "OPS" Or SNV_RS("TXN_TYP") = "AIN" Or SNV_RS("TXN_TYP") = "SAL" Or SNV_RS("TXN_TYP") = "PRT" Or SNV_RS("TXN_TYP") = "STO" Then
                CLD_RS("NUM_ONE") = CLD_RS("NUM_ONE") + SNV_RS("R_SHL_NOS") + SNV_RS("BRK_SHL")
                CLD_RS("NUM_TWO") = CLD_RS("NUM_TWO") + SNV_RS("R_PCK_NOS") + SNV_RS("BRK_PCK")
                CLD_RS.Update
            ElseIf SNV_RS("TXN_TYP") = "AOU" Or SNV_RS("TXN_TYP") = "PUR" Or SNV_RS("TXN_TYP") = "SRN" Or SNV_RS("TXN_TYP") = "STI" Then
                CLD_RS("NUM_ONE") = CLD_RS("NUM_ONE") - SNV_RS("R_SHL_NOS") + SNV_RS("BRK_SHL")
                CLD_RS("NUM_TWO") = CLD_RS("NUM_TWO") - SNV_RS("R_PCK_NOS") + SNV_RS("BRK_PCK")
                CLD_RS.Update
            End If
        Else
            CLD_RS.AddNew
            CLD_RS("CHR_TEN") = SNV_RS("PRD_PCK")
            CLD_RS("NUM_ONE") = 0
            CLD_RS("NUM_TWO") = 0
            CLD_RS("NUM_THR") = SNV_RS("PRD_UPC")
            If SNV_RS("TXN_TYP") = "OPS" Or SNV_RS("TXN_TYP") = "AIN" Or SNV_RS("TXN_TYP") = "SAL" Or SNV_RS("TXN_TYP") = "PRT" Or SNV_RS("TXN_TYP") = "STO" Then
                CLD_RS("NUM_ONE") = CLD_RS("NUM_ONE") + SNV_RS("R_SHL_NOS") + SNV_RS("BRK_SHL")
                CLD_RS("NUM_TWO") = CLD_RS("NUM_TWO") + SNV_RS("R_PCK_NOS") + SNV_RS("BRK_PCK")
                CLD_RS.Update
            ElseIf SNV_RS("TXN_TYP") = "AOU" Or SNV_RS("TXN_TYP") = "PUR" Or SNV_RS("TXN_TYP") = "SRN" Or SNV_RS("TXN_TYP") = "STI" Then
                CLD_RS("NUM_ONE") = CLD_RS("NUM_ONE") - SNV_RS("R_SHL_NOS") + SNV_RS("BRK_SHL")
                CLD_RS("NUM_TWO") = CLD_RS("NUM_TWO") - SNV_RS("R_PCK_NOS") + SNV_RS("BRK_PCK")
                CLD_RS.Update
            End If
            CLD_RS.Update
        End If
        SNV_RS.MoveNext
    Wend
End Sub
Private Sub ACTPAYDEX(DEN_DATE As Date)
EN_DAT = DEN_DATE
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Dim rs6 As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
con.Execute "update opa set opa_dbt=0", d
con.Execute "update opa set opa_crt=0", d
con.Execute "DELETE FROM OPA", d
con.Execute "UPDATE DIR SET DIR_MSG='" & Format(EN_DATE, "DD/MM/YYYY") & "' WHERE DIR_IDY = 'HDR' AND DIR_IDX <> 'DAT'"

rs.Open "select sum(txn_amt),dbt_idy from vch where txn_dat <=ctod('" & Format(EN_DAT, "MM/DD/YYYY") & "') group by dbt_idy", con, adOpenKeyset, adLockPessimistic
rs2.Open "select sum(txn_amt),crt_idy from vch where txn_dat <=ctod('" & Format(EN_DAT, "MM/DD/YYYY") & "') group by crt_idy", con, adOpenKeyset, adLockPessimistic

rs3.Open "select sum(NET_amt),CST_idy from INL where INV_dat <=ctod('" & Format(EN_DAT, "MM/DD/YYYY") & "') AND  txn_typ in('LSL','SAL','PRT','RSI','STO') group by CST_idy", con, adOpenKeyset, adLockPessimistic
rs4.Open "select sum(NET_amt),ACT_idy from INL where INV_dat <= ctod('" & Format(EN_DAT, "MM/DD/YYYY") & "') AND txn_typ in('LSL','SAL','PRT','RSI','STO') group by ACT_idy", con, adOpenKeyset, adLockPessimistic

rs5.Open "select sum(net_amt),cst_idy from inl where inv_dat <= ctod('" & EN_DAT & "') and txn_typ in('PUR','SRN','RSR','STI','OPS') group by cst_idy", con, adOpenKeyset, adLockPessimistic
rs6.Open "select sum(net_amt),act_idy from inl where inv_dat <= ctod('" & EN_DAT & "') and txn_typ in('PUR','SRN','RSR','STI','OPS') group by act_idy", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    con.Execute "update opa set opa_dbt = opa_dbt+ " & Val(rs(0)) & " where act_idy = '" & Trim(rs!dbt_idy) & "'", d
    If d = 0 Then con.Execute "insert into opa values('" & Trim(rs!dbt_idy) & "'," & Val(rs(0)) & ",0,ctod('" & TO_DAT & "'))"
    rs.MoveNext
Next
For A = 1 To rs2.RecordCount
    con.Execute "update opa set opa_crt = opa_crt+ " & Val(rs2(0)) & " where act_idy = '" & Trim(rs2!crt_idy) & "'", d
    If d = 0 Then con.Execute "insert into opa values('" & Trim(rs2!crt_idy) & "',0," & Val(rs2(0)) & ",ctod('" & TO_DAT & "'))"
    rs2.MoveNext
Next
For A = 1 To rs3.RecordCount
    con.Execute "update opa set opa_dbt = opa_dbt+ " & Val(rs3(0)) & " where act_idy = '" & Trim(rs3!CST_IDY) & "'", d
    If d = 0 Then con.Execute "insert into opa values('" & Trim(rs3!CST_IDY) & "'," & Val(rs3(0)) & ",0,ctod('" & TO_DAT & "'))"
    rs3.MoveNext
Next
For A = 1 To rs4.RecordCount
    con.Execute "update opa set opa_crt = opa_crt+ " & Val(rs4(0)) & " where act_idy = '" & Trim(rs4!ACT_IDY) & "'", d
    If d = 0 Then con.Execute "insert into opa values('" & Trim(rs4!ACT_IDY) & "',0," & Val(rs4(0)) & ",ctod('" & TO_DAT & "'))"
    rs4.MoveNext
Next
For A = 1 To rs6.RecordCount
    con.Execute "update opa set opa_dbt = opa_dbt+ " & Val(rs6(0)) & " where act_idy = '" & Trim(rs6!ACT_IDY) & "'", d
    If d = 0 Then con.Execute "insert into opa values('" & Trim(rs6!ACT_IDY) & "'," & Val(rs6(0)) & ",0,ctod('" & TO_DAT & "'))"
    rs6.MoveNext
Next
For A = 1 To rs5.RecordCount
    con.Execute "update opa set opa_crt = opa_crt+ " & Val(rs5(0)) & " where act_idy = '" & Trim(rs5!CST_IDY) & "'", d
    If d = 0 Then con.Execute "insert into opa values('" & Trim(rs5!CST_IDY) & "',0," & Val(rs5(0)) & ",ctod('" & TO_DAT & "'))"
    rs5.MoveNext
Next
End Sub
Private Sub COLGODDEX(DEN_DATE As Date)
EN_DAT = DEN_DATE
Dim SNV_RS As New ADODB.Recordset
Dim CLD_RS As New ADODB.Recordset
    If SNV_RS.State = 1 Then SNV_RS.Close
    If CLD_RS.State = 1 Then CLD_RS.Close
    con.Execute "DELETE FROM RP1"
    CLD_RS.Open "SELECT * FROM RP1", con, adOpenKeyset, adLockPessimistic
    SNV_RS.Open "SELECT * FROM SNV WHERE TXN_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    If SNV_RS.RecordCount > 0 Then SNV_RS.MoveFirst
    While Not SNV_RS.EOF
        If CLD_RS.State = 1 Then CLD_RS.Close
        CLD_RS.Open "SELECT * FROM RP1 WHERE CHR_TEN = '" & Trim(SNV_RS("PRD_PCK")) & "'", con, adOpenKeyset, adLockPessimistic
        CLD_RS.Find "CHR_TEN='" & Trim(SNV_RS("PRD_PCK")) & "'"
        If CLD_RS.EOF = False Then
            If SNV_RS("TXN_TYP") = "PMT" Or SNV_RS("TXN_TYP") = "AIN" Or SNV_RS("TXN_TYP") = "PUR" Then
                CLD_RS("NUM_ONE") = CLD_RS("NUM_ONE") - SNV_RS("D_SHL_NOS")
                CLD_RS("NUM_TWO") = CLD_RS("NUM_TWO") - SNV_RS("D_PCK_NOS")
                CLD_RS.Update
            ElseIf SNV_RS("TXN_TYP") = "AOU" Or SNV_RS("TXN_TYP") = "PRT" Or SNV_RS("TXN_TYP") = "STO" Then
                CLD_RS("NUM_ONE") = CLD_RS("NUM_ONE") + SNV_RS("D_SHL_NOS")
                CLD_RS("NUM_TWO") = CLD_RS("NUM_TWO") + SNV_RS("D_PCK_NOS")
                CLD_RS.Update
            ElseIf SNV_RS("TXN_TYP") = "DPC" Or SNV_RS("TXN_TYP") = "AIN" Or SNV_RS("TXN_TYP") = "PUR" Then
                CLD_RS("NUM_FIV") = CLD_RS("NUM_FIV") + SNV_RS("DP_SHL_NOS")
                CLD_RS("NUM_SIX") = CLD_RS("NUM_SIX") + SNV_RS("DP_PCK_NOS")
                CLD_RS.Update
            ElseIf SNV_RS("TXN_TYP") = "DPI" Or SNV_RS("TXN_TYP") = "PRT" Or SNV_RS("TXN_TYP") = "STO" Or SNV_RS("TXN_TYP") = "AOU" Then
                CLD_RS("NUM_FIV") = CLD_RS("NUM_FIV") - SNV_RS("DP_SHL_NOS")
                CLD_RS("NUM_SIX") = CLD_RS("NUM_SIX") - SNV_RS("DP_PCK_NOS")
                CLD_RS.Update
            End If
        Else
            CLD_RS.AddNew
            CLD_RS("CHR_TEN") = SNV_RS("PRD_PCK")
            CLD_RS("NUM_ONE") = 0
            CLD_RS("NUM_TWO") = 0
            CLD_RS("NUM_THR") = SNV_RS("PRD_UPC")
            CLD_RS("NUM_FIV") = 0
            CLD_RS("NUM_SIX") = 0
            If SNV_RS("TXN_TYP") = "PMT" Or SNV_RS("TXN_TYP") = "AIN" Or SNV_RS("TXN_TYP") = "PUR" Then
                CLD_RS("NUM_ONE") = CLD_RS("NUM_ONE") - SNV_RS("D_SHL_NOS")
                CLD_RS("NUM_TWO") = CLD_RS("NUM_TWO") - SNV_RS("D_PCK_NOS")
                CLD_RS.Update
            ElseIf SNV_RS("TXN_TYP") = "AOU" Or SNV_RS("TXN_TYP") = "PRT" Or SNV_RS("TXN_TYP") = "STO" Then
                CLD_RS("NUM_ONE") = CLD_RS("NUM_ONE") + SNV_RS("D_SHL_NOS")
                CLD_RS("NUM_TWO") = CLD_RS("NUM_TWO") + SNV_RS("D_PCK_NOS")
                CLD_RS.Update
            ElseIf SNV_RS("TXN_TYP") = "DPC" Or SNV_RS("TXN_TYP") = "AIN" Or SNV_RS("TXN_TYP") = "PUR" Then
                CLD_RS("NUM_FIV") = CLD_RS("NUM_FIV") + SNV_RS("DP_SHL_NOS")
                CLD_RS("NUM_SIX") = CLD_RS("NUM_SIX") + SNV_RS("DP_PCK_NOS")
                CLD_RS.Update
            ElseIf SNV_RS("TXN_TYP") = "DPI" Or SNV_RS("TXN_TYP") = "PRT" Or SNV_RS("TXN_TYP") = "STO" Or SNV_RS("TXN_TYP") = "AOU" Then
                CLD_RS("NUM_FIV") = CLD_RS("NUM_FIV") - SNV_RS("DP_SHL_NOS")
                CLD_RS("NUM_SIX") = CLD_RS("NUM_SIX") - SNV_RS("DP_PCK_NOS")
                CLD_RS.Update
            End If
            CLD_RS.Update
        End If
        SNV_RS.MoveNext
    Wend
End Sub
Private Sub STKTXTDEX(DEN_DATE As Date)
    EN_DAT = DEN_DATE
    If rs.State = 1 Then rs.Close
    con.Execute "DELETE FROM CLD"
    rs.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockPessimistic
    If Me.Tag = "COK_STK_BAT" Or Me.Tag = "COK_STK_BRD" Or Me.Tag = "COK_STK_SEG" Then
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where inv_dat < ctod('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika and Ramanesh
    Else
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat < ctod('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic
    End If
    For A = 1 To RS1.RecordCount
        If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!sal_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!sal_stk = Val(RS1(0)) * -1
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "OPS" Then
            con.Execute "UPDATE CLD SET OPN_STK = OPN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!opn_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "PUR" Then
        con.Execute "UPDATE CLD SET PUR_STK = PUR_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!PUR_STK = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "PRT" Then
        con.Execute "UPDATE CLD SET PRT_STK = PRT_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!PRT_STK = Val(RS1(0))
                rs!PUR_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "STI" Then
        con.Execute "UPDATE CLD SET STI_STK = STI_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!STI_STK = Val(RS1(0))
                rs!PRT_STK = 0
                rs!PUR_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "SRN" Then
        con.Execute "UPDATE CLD SET SRN_STK = SRN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "STO" Or RS1!TXN_TYP = "ADJ" Then
        con.Execute "UPDATE CLD SET STO_STK = STO_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs(7) = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs!PUR_STK = 0
                rs!cls_stk = 0
                rs.Update
            End If
        End If
        RS1.MoveNext
    Next
    rs.Close
    con.Execute "UPDATE CLD SET CLS_STK=(PUR_STK+OPN_STK+SRN_STK+STI_STK)-(SAL_STK+PRT_STK+STO_STK)"
    con.Execute "UPDATE CLD SET OPN_STK=CLS_STK"
    con.Execute "UPDATE CLD SET PUR_STK=0"
    con.Execute "UPDATE CLD SET PRT_STK=0"
    con.Execute "UPDATE CLD SET SAL_STK=0"
    con.Execute "UPDATE CLD SET STO_STK=0"
    con.Execute "UPDATE CLD SET SRN_STK=0"
    con.Execute "UPDATE CLD SET STI_STK=0"
    con.Execute "UPDATE CLD SET CLS_STK=0"
    
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockPessimistic
    If Me.Tag = "COK_STK_BAT" Or Me.Tag = "COK_STK_BRD" Or Me.Tag = "COK_STK_SEG" Then
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where inv_dat >= ctod('" & EN_DAT & "') AND INV_DAT <=CTOD('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika & Ramanesh
    Else
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat >= ctod('" & EN_DAT & "') AND INV_DAT <=CTOD('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika & Ramanesh
    End If
    For A = 1 To RS1.RecordCount
        If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!sal_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!sal_stk = Val(RS1(0)) * -1
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "OPS" Then
            con.Execute "UPDATE CLD SET OPN_STK = OPN_STK +" & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!opn_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "PUR" Then
        con.Execute "UPDATE CLD SET PUR_STK = PUR_STK +" & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!PUR_STK = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "PRT" Then
        con.Execute "UPDATE CLD SET PRT_STK = PRT_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!PRT_STK = Val(RS1(0))
                rs!PUR_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "STI" Then
        con.Execute "UPDATE CLD SET STI_STK = STI_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!STI_STK = Val(RS1(0))
                rs!PRT_STK = 0
                rs!PUR_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "SRN" Then
        con.Execute "UPDATE CLD SET SRN_STK = SRN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "STO" Or RS1!TXN_TYP = "ADJ" Then
        con.Execute "UPDATE CLD SET STO_STK = STO_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs(7) = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs!PUR_STK = 0
                rs!cls_stk = 0
                rs.Update
            End If
        End If
        RS1.MoveNext
    Next
    rs.Close
    con.Execute "UPDATE CLD SET CLS_STK=(PUR_STK+OPN_STK+SRN_STK+STI_STK)-(SAL_STK+PRT_STK+STO_STK)"
End Sub
Private Sub CHN_MEX()
Dim rs As New ADODB.Recordset
Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim STR_VAL As Integer
Dim TOT_UPD As Integer

Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\CHN.TXT") = True Then
    FILE_NAME = "C:\DEX\CHN.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("Master Data Xchange for Channels Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("Master Data Xchange for Channels Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim CHN_NME As String
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\CHN.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        CHN_NME = ""
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        If Mid(INL_STR, STR_VAL, STR_LEN) = "CHN" Then
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> ";" Then
                        CHN_NME = CHN_NME + Mid(INL_STR, J, 1)
                    Else
                        Exit For
                    End If
                Next J
        Else
            MsgBox ("Invalid Channel Record Encountered!, Contact effmcg@consoftsystems.com !"), vbCritical, "DAS Version 4.0"
            MsgBox (TOT_UPD & " Channels Updated in Channel Master "), vbInformation, "DAS Version 4.0"
            Exit Sub
        End If
        If rs.State = 1 Then rs.Close
        rs.Open "select * from CAT where alltrim(CAT_NME) = '" & UCase(Trim(CHN_NME)) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.EOF() = True Then
            rs.AddNew
            rs("CAT_IDY") = increment("IDY_CAT")
            rs("CAT_NME") = UCase(Trim(CHN_NME))
            rs.Update
            TOT_UPD = TOT_UPD + 1
        End If
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_UPD & " Channels Updated in Channel Master "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\CHN.TXT") = True Then
        FS_SOU.DeleteFile ("C:\DEX\CHN.TXT")
    End If
End Sub
Private Sub RTT_MEX()
Dim rs As New ADODB.Recordset
Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim TOT_UPD As Integer
Dim STR_VAL As Integer
Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\RTT.TXT") = True Then
    FILE_NAME = "C:\DEX\RTT.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("Master Data Xchange for Retailer Types Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("Master Data Xchange for Retailer Types Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim RTT_NME As String
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\RTT.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        RTT_NME = ""
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        If Mid(INL_STR, STR_VAL, STR_LEN) = "RTT" Then
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> ";" Then
                        RTT_NME = RTT_NME + Mid(INL_STR, J, 1)
                    Else
                        Exit For
                    End If
                Next J
        Else
            MsgBox ("Invalid Retailer Type Record Encountered!, Contact effmcg@consoftsystems.com !"), vbCritical, "DAS Version 4.0"
            MsgBox (TOT_UPD & " Retailer Types Updated in Retailer Type Master "), vbInformation, "DAS Version 4.0"
            Exit Sub
        End If
        If rs.State = 1 Then rs.Close
        rs.Open "select * from DIR where alltrim(DIR_ALS) = '" & UCase(Trim(RTT_NME)) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.EOF() = True Then
            rs.AddNew
            rs("DIR_IDY") = "RET"
            rs("DIR_IDX") = "TYA"
            rs("DIR_ALS") = UCase(RTT_NME)
            rs("DIR_MSG") = UCase(RTT_NME)
            rs.Update
            TOT_UPD = TOT_UPD + 1
        End If
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_UPD & " Retailer Types Updated in Retailer Type Master "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\RTT.TXT") = True Then
        FS_SOU.DeleteFile ("C:\DEX\RTT.TXT")
    End If
End Sub
Private Sub REG_MEX()
Dim rs As New ADODB.Recordset
Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim TOT_UPD As Integer
Dim STR_VAL As Integer
Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\REG.TXT") = True Then
    FILE_NAME = "C:\DEX\REG.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("Master Data Xchange for Regions Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("Master Data Xchange for Regions Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim REG_NME As String
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\REG.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        REG_NME = ""
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        If Mid(INL_STR, STR_VAL, STR_LEN) = "REG" Then
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> ";" Then
                        REG_NME = REG_NME + Mid(INL_STR, J, 1)
                    Else
                        Exit For
                    End If
                Next J
        Else
            MsgBox ("Invalid Region Record Encountered!, Contact effmcg@consoftsystems.com !"), vbCritical, "DAS Version 4.0"
            MsgBox (TOT_UPD & " Regions Updated in Region Master "), vbInformation, "DAS Version 4.0"
            Exit Sub
        End If
        If rs.State = 1 Then rs.Close
        rs.Open "select * from DIR where alltrim(DIR_ALS) = '" & UCase(Trim(REG_NME)) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.EOF() = True Then
            rs.AddNew
            rs("DIR_IDY") = "CCI"
            rs("DIR_IDX") = "RGA"
            rs("DIR_ALS") = UCase(REG_NME)
            rs("DIR_MSG") = UCase(REG_NME)
            rs.Update
            TOT_UPD = TOT_UPD + 1
        End If
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_UPD & " Regions Updated in Region Master "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\REG.TXT") = True Then
        FS_SOU.DeleteFile ("C:\DEX\REG.TXT")
    End If
End Sub

Private Sub DST_MEX()
Dim rs As New ADODB.Recordset
Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim TOT_UPD As Integer
Dim STR_VAL As Integer
Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\DST.TXT") = True Then
    FILE_NAME = "C:\DEX\DST.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("Master Data Xchange for District-States Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("Master Data Xchange for District-States Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim M_CTR As Integer
    Dim FLD_NME As String
    Dim DIS_NME As String
    Dim STA_NME As String
    Dim DIS_POP As String
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\DST.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        FLD_NME = ""
        DIS_NME = ""
        STA_NME = ""
        DIS_POP = ""
        M_CTR = 1
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        If Mid(INL_STR, STR_VAL, STR_LEN) = "DST" Then
            For M_CTR = 1 To 3
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> ";" Then
                        FLD_NME = FLD_NME + Mid(INL_STR, J, 1)
                    Else
                        If M_CTR = 1 Then
                            DIS_NME = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 2 Then
                            STA_NME = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 3 Then
                            DIS_POP = FLD_NME
                            FLD_NME = ""
                        End If
                        M_CTR = M_CTR + 1
                    End If
                Next J
            Next M_CTR
        Else
            MsgBox ("Invalid District-State Record Encountered!, Contact effmcg@consoftsystems.com !"), vbCritical, "DAS Version 4.0"
            MsgBox (TOT_UPD & " District-States Updated in District-State Master "), vbInformation, "DAS Version 4.0"
            Exit Sub
        End If
        If rs.State = 1 Then rs.Close
        rs.Open "select * from DIR where alltrim(DIR_MSG) = '" & UCase(Trim(DIS_NME)) & "," & UCase(Trim(STA_NME)) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.EOF() = True Then
            rs.AddNew
            rs("DIR_IDY") = "STA"
            rs("DIR_IDX") = "DIS"
            rs("DIR_ALS") = UCase(DIS_POP)
            rs("DIR_MSG") = UCase(DIS_NME) & "," & UCase(STA_NME)
            rs.Update
            TOT_UPD = TOT_UPD + 1
        End If
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_UPD & " District-States Updated in District-State Master "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\DST.TXT") = True Then
        FS_SOU.DeleteFile ("C:\DEX\DST.TXT")
    End If
End Sub
Private Sub KAC_MEX()
Dim rs As New ADODB.Recordset
Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim TOT_UPD As Integer
Dim STR_VAL As Integer
Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\KAC.TXT") = True Then
    FILE_NAME = "C:\DEX\KAC.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("Master Data Xchange for Key Accounts Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("Master Data Xchange for Key Accounts Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim KAC_NME As String
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\KAC.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        KAC_NME = ""
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        If Mid(INL_STR, STR_VAL, STR_LEN) = "KAC" Then
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> ";" Then
                        KAC_NME = KAC_NME + Mid(INL_STR, J, 1)
                    Else
                        Exit For
                    End If
                Next J
        Else
            MsgBox ("Invalid Key Account Record Encountered!, Contact effmcg@consoftsystems.com !"), vbCritical, "DAS Version 4.0"
            MsgBox (TOT_UPD & " Key Accounts Updated in Key Account Master "), vbInformation, "DAS Version 4.0"
            Exit Sub
        End If
        If rs.State = 1 Then rs.Close
        rs.Open "select * from DIR where alltrim(DIR_ALS) = '" & UCase(Trim(KAC_NME)) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.EOF() = True Then
            rs.AddNew
            rs("DIR_IDY") = "KAC"
            rs("DIR_IDX") = "NKA"
            rs("DIR_ALS") = UCase(KAC_NME)
            rs("DIR_MSG") = UCase(KAC_NME)
            rs.Update
            TOT_UPD = TOT_UPD + 1
        End If
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_UPD & " Key Accounts Updated in Key Account Master "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\KAC.TXT") = True Then
        FS_SOU.DeleteFile ("C:\DEX\KAC.TXT")
    End If
End Sub
Private Sub KOC_MEX()
Dim rs As New ADODB.Recordset
Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim TOT_UPD As Integer
Dim STR_VAL As Integer
Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\KOC.TXT") = True Then
    FILE_NAME = "C:\DEX\KOC.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("Master Data Xchange for Sales Calendar Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("Master Data Xchange for Sales Calendar Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim M_CTR As Integer
    Dim FLD_NME As String
    Dim KO_WEK As String
    Dim FRM_DAT As String
    Dim TO_DAT As String
    Dim KO_MTH As String
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\KOC.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        FLD_NME = ""
        KO_WEK = ""
        FRM_DAT = ""
        TO_DAT = ""
        M_CTR = 1
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        If Mid(INL_STR, STR_VAL, STR_LEN) = "KOC" Then
            For M_CTR = 1 To 4
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> ";" Then
                        FLD_NME = FLD_NME + Mid(INL_STR, J, 1)
                    Else
                        If M_CTR = 1 Then
                            KO_WEK = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 2 Then
                            FRM_DAT = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 3 Then
                            TO_DAT = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 4 Then
                            KO_MTH = FLD_NME
                            FLD_NME = ""
                        End If
                        M_CTR = M_CTR + 1
                    End If
                Next J
            Next M_CTR
        Else
            MsgBox ("Invalid Sales Calendar Record Encountered!, Contact effmcg@consoftsystems.com !"), vbCritical, "DAS Version 4.0"
            MsgBox (M_CTR & " Entries in Sales Calendar Updated "), vbInformation, "DAS Version 4.0"
            Exit Sub
        End If
        If rs.State = 1 Then rs.Close
        rs.Open "select * from WEK where alltrim(str(WEK_IDY)) = '" & (Trim(KO_WEK)) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.EOF() = True Then
            rs.AddNew
            rs("WEK_IDY") = Val(Trim(KO_WEK))
            rs("FRM_DAT") = FRM_DAT
            rs("TO_DAT") = TO_DAT
            rs("MTH_NOS") = Val(Trim(KO_MTH))
            rs.Update
            TOT_UPD = TOT_UPD + 1
        End If
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_UPD & " Entries in Sales Calendar Updated "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\KOC.TXT") = True Then
        FS_SOU.DeleteFile ("C:\DEX\KOC.TXT")
    End If
End Sub
Private Sub DPF_MEX()
Dim rs As New ADODB.Recordset
Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim TOT_UPD As Integer
Dim STR_VAL As Integer
Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\DPF.TXT") = True Then
    FILE_NAME = "C:\DEX\DPF.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("Master Data Xchange for Distributor Profile Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("Master Data Xchange for Distributor Profile Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim M_CTR As Integer
    Dim FLD_NME As String
    
    Dim M_DIS_NME As String
    Dim M_DIS_AD1 As String
    Dim M_DIS_AD2 As String
    Dim M_DIS_TIN As String
    Dim M_DIS_GST As String
    Dim M_DIS_CST As String
    Dim M_DIS_REG As String
    Dim M_DIS_DEP As String
    Dim M_DIS_UCO As String
    Dim M_DIS_TWN As String
    Dim M_DIS_IDY As String
    
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\DPF.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        FLD_NME = ""
        M_CTR = 1
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        If Mid(INL_STR, STR_VAL, STR_LEN) = "DPF" Then
            For M_CTR = 1 To 11
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> ";" Then
                        FLD_NME = FLD_NME + Mid(INL_STR, J, 1)
                    Else
                        If M_CTR = 1 Then
                            M_DIS_NME = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 2 Then
                            M_DIS_AD1 = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 3 Then
                            M_DIS_AD2 = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 4 Then
                            M_DIS_TIN = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 5 Then
                            M_DIS_GST = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 6 Then
                            M_DIS_CST = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 7 Then
                            M_DIS_REG = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 8 Then
                            M_DIS_DEP = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 9 Then
                            M_DIS_UCO = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 10 Then
                            M_DIS_TWN = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 11 Then
                            M_DIS_IDY = FLD_NME
                            FLD_NME = ""
                        End If
                        M_CTR = M_CTR + 1
                        If rs.State = 1 Then rs.Close
                    rs.Open "select * from dir WHERE DIR_IDX = 'TT1'", con, adOpenKeyset, adLockPessimistic
                    If rs.RecordCount > 0 Then rs("DIR_MSG") = UCase(M_DIS_NME) & ""
                    rs.Update
                    If rs.State = 1 Then rs.Close
                    rs.Open "select * from dir WHERE DIR_IDX = 'TT2'", con, adOpenKeyset, adLockPessimistic
                    If rs.RecordCount > 0 Then rs("DIR_MSG") = UCase(M_DIS_AD1) & ""
                    rs.Update
                    If rs.State = 1 Then rs.Close
                    rs.Open "select * from dir WHERE DIR_IDX = 'TT3'", con, adOpenKeyset, adLockPessimistic
                    If rs.RecordCount > 0 Then rs("DIR_MSG") = UCase(M_DIS_AD2) & ""
                    rs.Update
                    If rs.State = 1 Then rs.Close
                    rs.Open "select * from dir WHERE DIR_IDX = 'GST'", con, adOpenKeyset, adLockPessimistic
                    If rs.RecordCount > 0 Then rs("DIR_MSG") = UCase(M_DIS_TIN) & ""
                    rs.Update
                    If rs.State = 1 Then rs.Close
                    rs.Open "select * from dir WHERE DIR_IDX = 'CST'", con, adOpenKeyset, adLockPessimistic
                    If rs.RecordCount > 0 Then rs("DIR_MSG") = UCase(M_DIS_GST) & ""
                    rs.Update
                    If rs.State = 1 Then rs.Close
                    rs.Open "select * from dir WHERE DIR_IDX = 'TT4'", con, adOpenKeyset, adLockPessimistic
                    If rs.RecordCount > 0 Then rs("DIR_MSG") = UCase(M_DIS_CST) & ""
                    rs.Update
                    If rs.State = 1 Then rs.Close
                    rs.Open "select * from dir WHERE DIR_IDY = 'ZON'", con, adOpenKeyset, adLockPessimistic
                    If rs.RecordCount > 0 Then rs("DIR_ALS") = UCase(M_DIS_REG) & ""
                    rs.Update
                    If rs.State = 1 Then rs.Close
                    rs.Open "select * from dir WHERE DIR_IDY = 'DEP'", con, adOpenKeyset, adLockPessimistic
                    If rs.RecordCount > 0 Then rs("DIR_ALS") = UCase(M_DIS_DEP) & ""
                    rs.Update
                    If rs.State = 1 Then rs.Close
                    rs.Open "select * from dir WHERE DIR_IDY = 'UNT'", con, adOpenKeyset, adLockPessimistic
                    If rs.RecordCount > 0 Then rs("DIR_ALS") = UCase(M_DIS_UCO) & ""
                    rs.Update
                    If rs.State = 1 Then rs.Close
                    rs.Open "select * from dir WHERE DIR_IDY = 'TWN'", con, adOpenKeyset, adLockPessimistic
                    If rs.RecordCount > 0 Then rs("DIR_ALS") = UCase(M_DIS_TWN) & ""
                    rs.Update
                    If rs.State = 1 Then rs.Close
                    rs.Open "select * from dir WHERE DIR_IDY = 'WDS'", con, adOpenKeyset, adLockPessimistic
                    If rs.RecordCount > 0 Then rs("DIR_ALS") = UCase(M_DIS_IDY) & ""
                    rs.Update
                    End If
                Next J
            Next M_CTR
        Else
            MsgBox ("Invalid Distributor Profile Data File & Record!, Contact effmcg@consoftsystems.com !"), vbCritical, "DAS Version 4.0"
        End If
        
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_REC & " Entries in Distributor Profile Updated "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\DPF.TXT") = True Then
        FS_SOU.DeleteFile ("C:\DEX\DPF.TXT")
    End If
End Sub
Private Sub PRD_MEX()

Dim rs As New ADODB.Recordset
Dim Pgr_RS As New ADODB.Recordset
Dim BAT_RS As New ADODB.Recordset
Dim ACT_RS As New ADODB.Recordset

Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim TOT_UPD As Integer
Dim STR_VAL As Integer

Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\PRD.TXT") = True Then
    FILE_NAME = "C:\DEX\PRD.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("Master Data Xchange for SKU - Product Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("Master Data Xchange for SKU - Product Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim M_CTR As Integer
    Dim FLD_NME As String
    
    Dim M_GRP_IDY As String
    Dim M_PRD_IDY As String
    Dim M_ACT_IDY As String
    
    Dim ACT_NME As String
    Dim GRP_CAT As String
    Dim GRP_CLS As String
    Dim PRD_NME As String
    Dim ALT_IDY As String
    Dim BRD_NME As String
    Dim PRD_PCK As String
    Dim PRD_PKT As String
    Dim PUR_RAT As String
    Dim PUR_TAX As String
    Dim PUR_DIS As String
    Dim SAL_RAT As String
    Dim SAL_TAX As String
    Dim SAL_DIS As String
    Dim MRP_RAT As String
    Dim VOL_WGT As String
    Dim PRD_UOM As String
    Dim SUB_UNT As Integer
    Dim TOD_PCG As String
    Dim LST_RAT As String
    Dim LST_DIS As String
    Dim SAF_LEV As String
    Dim SAF_DAY As Integer
    Dim stk_typ As String
    
    Dim STR_CMP As String
    
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\PRD.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        FLD_NME = ""
        M_CTR = 1
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        STR_CMP = Mid(Trim(INL_STR), STR_VAL, STR_LEN)
        If STR_CMP = "PRD" Then
            For M_CTR = 1 To 24
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> ";" Then
                        FLD_NME = FLD_NME + Mid(INL_STR, J, 1)
                    Else
                        If M_CTR = 1 Then
                            GRP_CAT = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 2 Then
                            GRP_CLS = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 3 Then
                            PRD_NME = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 4 Then
                            ALT_IDY = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 5 Then
                            BRD_NME = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 6 Then
                            PRD_PCK = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 7 Then
                            PRD_PKT = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 8 Then
                            PUR_RAT = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 9 Then
                            PUR_TAX = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 10 Then
                            PUR_DIS = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 11 Then
                            SAL_RAT = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 12 Then
                            SAL_TAX = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 13 Then
                            SAL_DIS = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 14 Then
                            MRP_RAT = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 15 Then
                            VOL_WGT = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 16 Then
                            PRD_UOM = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 17 Then
                            SUB_UNT = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        If M_CTR = 18 Then
                            ACT_NME = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 19 Then
                            SAF_LEV = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 20 Then
                            SAF_DAY = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        If M_CTR = 21 Then
                            stk_typ = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 22 Then
                            TOD_PCG = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 23 Then
                            LST_RAT = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        If M_CTR = 24 Then
                            LST_DIS = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        M_CTR = M_CTR + 1
                    End If
                Next J
            Next M_CTR
        Else
            MsgBox ("Invalid SKU - Product Record Encountered!, Contact effmcg@consoftsystems.com !"), vbCritical, "DAS Version 4.0"
            MsgBox (TOT_UPD & " Entries in SKU - Product Updated "), vbInformation, "DAS Version 4.0"
            Exit Sub
        End If
        If rs.State = 1 Then rs.Close
        rs.Open "select * from PRD where ALT_IDY = '" & Trim(ALT_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            rs("PRD_NME") = UCase(PRD_NME)
            rs("BRD_NME") = UCase(BRD_NME)
            rs("PRD_PCK") = UCase(PRD_PCK)
            rs("PRD_PKT") = UCase(PRD_PKT)
            rs("PUR_RAT") = Val(PUR_RAT)
            rs("PUR_TAX") = Val(PUR_TAX)
            rs("PUR_DIS") = Val(PUR_DIS)
            rs("SAL_RAT") = Val(SAL_RAT)
            rs("SAL_TAX") = Val(SAL_TAX)
            rs("SAL_DIS") = Val(SAL_DIS)
            rs("MRP_RAT") = Val(MRP_RAT)
            rs("CON_FAC") = Val(VOL_WGT)
            rs("PRD_UOM") = PRD_UOM
            rs("TOD_PCG") = Val(TOD_PCG)
            rs("LST_RAT") = Val(LST_RAT)
            rs("LST_DIS") = Val(LST_DIS)
            rs.Update
        Else
            If ACT_RS.State = 1 Then ACT_RS.Close
            ACT_RS.Open "select * from ACT where alltrim(ACT_NME) = '" & UCase(Trim(ACT_NME)) & "'", con, adOpenKeyset, adLockPessimistic
            If ACT_RS.RecordCount > 0 Then
                M_ACT_IDY = ACT_RS("ACT_IDY")
            Else
                ACT_RS.AddNew
                M_ACT_IDY = increment("IDY_ACT")
                ACT_RS("ACT_IDY") = M_ACT_IDY
                ACT_RS("GRP_IDY") = "0000000008"
                ACT_RS("act_nme") = UCase(ACT_NME)
                ACT_RS("act_ad1") = ""
                ACT_RS("act_ad2") = ""
                ACT_RS("act_ad3") = ""
                ACT_RS("act_ad4") = ""
                ACT_RS("act_zip") = ""
                ACT_RS("cst_typ") = "Retailer"
                ACT_RS("KEY_ACT") = "N"
                ACT_RS("ACT_TYP") = "PUR"
                ACT_RS("act_tfx") = ""
                ACT_RS("act_fax") = ""
                ACT_RS("act_mbl") = ""
                ACT_RS("act_pgr") = ""
                ACT_RS("act_eml") = ""
                ACT_RS("ACT_WWW") = ""
                ACT_RS("act_cpn") = ""
                ACT_RS("act_dsg") = ""
                ACT_RS("DBT_TYP") = "Y"
                ACT_RS("ACT_DOC") = Date
                ACT_RS("act_gst") = ""
                ACT_RS("act_cst") = ""
                ACT_RS("act_rgd") = ""
                ACT_RS("acn_one") = ""
                ACT_RS("bnm_one") = ""
                ACT_RS("bbn_one") = ""
                ACT_RS("acn_two") = ""
                ACT_RS("bnm_two") = ""
                ACT_RS("bbn_two") = ""
                ACT_RS("ret_typ") = ""
                ACT_RS("van_cap") = "0"
                ACT_RS("GRP_CLF") = "Creditor"
                ACT_RS("LOC_IDY") = ""
                ACT_RS("CAT_idy") = ""
                ACT_RS("mar_idy") = ""
                ACT_RS("smn_idy") = ""
                ACT_RS("rou_idy") = ""
                ACT_RS("zon_idy") = ""
                ACT_RS("dbt_DIs") = "0"
                ACT_RS("mcl_req") = "N"
                ACT_RS("dbt_mcl") = "0.00"
                ACT_RS("dbt_mcp") = "0.00"
                ACT_RS("dbt_rap") = ""
                ACT_RS("crt_dis") = "0.00"
                ACT_RS("mdl_req") = "N"
                ACT_RS("crt_mdl") = "0.00"
                ACT_RS("crt_ppd") = 0
                ACT_RS.Update
            End If
            
            If Pgr_RS.State = 1 Then Pgr_RS.Close
            Pgr_RS.Open "select * from PGR where GRP_CAT = '" & UCase(GRP_CAT) & "' and GRP_CLS = '" & UCase(GRP_CLS) & "'", con, adOpenKeyset, adLockPessimistic
            If Pgr_RS.RecordCount > 0 Then
                M_GRP_IDY = Pgr_RS("GRP_IDY")
            Else
                Pgr_RS.AddNew
                Pgr_RS("GRP_CAT") = UCase(GRP_CAT)
                Pgr_RS("GRP_CLS") = UCase(GRP_CLS)
                M_GRP_IDY = increment("IDY_PGR")
                Pgr_RS("GRP_IDY") = M_GRP_IDY
                Pgr_RS.Update
            End If
            
            rs.AddNew
            rs("GRP_IDY") = M_GRP_IDY
            M_PRD_IDY = increment("IDY_PRD")
            rs("PRD_IDY") = M_PRD_IDY
            rs("ALT_IDY") = ALT_IDY
            rs("PRD_PRN") = ALT_IDY
            rs("PRD_NME") = UCase(PRD_NME)
            rs("BRD_NME") = UCase(BRD_NME)
            rs("PRD_PCK") = UCase(PRD_PCK)
            rs("PRD_PKT") = UCase(PRD_PKT)
            rs("PUR_RAT") = Val(PUR_RAT)
            rs("PUR_TAX") = Val(PUR_TAX)
            rs("PUR_DIS") = Val(PUR_DIS)
            rs("SAL_RAT") = Val(SAL_RAT)
            rs("SAL_TAX") = Val(SAL_TAX)
            rs("SAL_DIS") = Val(SAL_DIS)
            rs("MRP_RAT") = Val(MRP_RAT)
            rs("CON_FAC") = Val(VOL_WGT)
            rs("PRD_UOM") = PRD_UOM
            rs("SUB_UNT") = SUB_UNT
            rs("CMP_IDY") = M_ACT_IDY 'Supplier ID Information
            rs("SAF_LEV") = Val(SAF_LEV)
            'rs("TOD_PCG") = SAF_DAY
            rs("TOD_PCG") = Val(TOD_PCG)
            rs("LST_RAT") = Val(LST_RAT)
            rs("LST_DIS") = Val(LST_DIS)
            rs("STK_TYP") = stk_typ
            rs("EXC_INF") = "N"
            rs("CUR_LEV") = "0.000"
            rs.Update
            TOT_UPD = TOT_UPD + 1
            
            If BAT_RS.State = 1 Then BAT_RS.Close
            BAT_RS.Open "select * from BRD", con, adOpenKeyset, adLockPessimistic
            
            BAT_RS.AddNew
            BAT_RS("PRD_IDY") = M_PRD_IDY
            BAT_RS("BAT_IDY") = "BAT" & T7increment("IDY_BAT")
            'BAT_RS("BAT_NME") = "MRP " & convert(Val(MRP_RAT) / SUB_UNT) & " " & Trim(Year(Date)) & "-" & Trim(Month(Date)) & "-" & Trim(Day(Date))
            If (Val(MRP_RAT) / SUB_UNT) = 0 Then
                BAT_RS("BAT_NME") = "MRP:" & "-" & " " & Format(Date, "YY") & Format(Date, "MM") & Format(Date, "dd")
            Else
                BAT_RS("BAT_NME") = "MRP:" & Trim(str((Val(MRP_RAT) / SUB_UNT))) & " " & Format(Date, "YY") & Format(Date, "MM") & Format(Date, "dd")
            End If
            BAT_RS("PRD_PDR") = Val(PUR_RAT)
            BAT_RS("PRD_SDR") = Val(SAL_RAT)
            BAT_RS("PRD_MRP") = Val(MRP_RAT)
            BAT_RS("PRD_QTY") = "0.000"
            BAT_RS("BRD_DOM") = Date
            BAT_RS("BRD_DOE") = DateAdd("D", Val(TOD_PCG), Date)
            BAT_RS("P_CRT_RNT") = Val(LST_RAT)
            BAT_RS("P_CRT_VAT") = Val(LST_DIS)
            BAT_RS("S_CRT_RNT") = Val(LST_RAT)
            BAT_RS("S_CRT_VAT") = Val(LST_DIS)
            BAT_RS("B_CRT_RAT") = Val(LST_RAT)
            BAT_RS("B_CRT_DIS") = Val(LST_DIS)
            BAT_RS("P_INC_DIS") = 0#
            BAT_RS("S_INC_DIS") = 0#
            BAT_RS.Update

        End If
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_UPD & " Entries in SKU - Product Updated "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\PRD.TXT") = True Then
        FS_SOU.DeleteFile ("C:\DEX\PRD.TXT")
    End If
End Sub
Private Sub BAT_MEX()

Dim rs As New ADODB.Recordset
Dim Pgr_RS As New ADODB.Recordset
Dim BAT_RS As New ADODB.Recordset
Dim ACT_RS As New ADODB.Recordset

Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim TOT_UPD As Integer
Dim STR_VAL As Integer

Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\BAT.TXT") = True Then
    FILE_NAME = "C:\DEX\BAT.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("Master Data Xchange for SKU - Batch Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("Master Data Xchange for SKU - Batch Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim M_CTR As Integer
    Dim FLD_NME As String
    
    Dim M_GRP_IDY As String
    Dim M_PRD_IDY As String
    Dim M_ACT_IDY As String
    
    Dim ACT_NME As String
    Dim GRP_CAT As String
    Dim GRP_CLS As String
    Dim PRD_NME As String
    Dim ALT_IDY As String
    Dim BRD_NME As String
    Dim PRD_PCK As String
    Dim PRD_PKT As String
    Dim PUR_RAT As String
    Dim PUR_TAX As String
    Dim PUR_DIS As String
    Dim SAL_RAT As String
    Dim SAL_TAX As String
    Dim SAL_DIS As String
    Dim MRP_RAT As String
    Dim VOL_WGT As String
    Dim PRD_UOM As String
    Dim LST_RAT As String
    Dim LST_DIS As String
    Dim PUR_INC As String
    Dim SAL_INC As String
    Dim SUB_UNT As Integer
    Dim SAF_LEV As Integer
    Dim SAF_DAY As Integer
    Dim stk_typ As String
    Dim MFG_DAT As Date
    Dim BBD_DAT As Date
    Dim Bat_nme As String
    
    
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\BAT.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        FLD_NME = ""
        M_CTR = 1
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        If Mid(INL_STR, STR_VAL, STR_LEN) = "BAT" Then
            For M_CTR = 1 To 15
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> ";" Then
                        FLD_NME = FLD_NME + Mid(INL_STR, J, 1)
                    Else
                        If M_CTR = 1 Then
                            ALT_IDY = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 2 Then
                            PUR_RAT = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 3 Then
                            SAL_RAT = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 4 Then
                            MRP_RAT = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 5 Then
                            MFG_DAT = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 6 Then
                            BBD_DAT = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 7 Then
                            Bat_nme = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 8 Then
                            PUR_TAX = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 9 Then
                            SAL_TAX = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 10 Then
                            PUR_DIS = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 11 Then
                            SAL_DIS = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 12 Then
                            LST_RAT = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        If M_CTR = 13 Then
                            LST_DIS = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        If M_CTR = 14 Then
                            PUR_INC = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        If M_CTR = 15 Then
                            SAL_INC = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        M_CTR = M_CTR + 1
                    End If
                Next J
            Next M_CTR
        Else
            MsgBox ("Invalid SKU - Batch Record Encountered!, Contact effmcg@consoftsystems.com !"), vbCritical, "DAS Version 4.0"
            MsgBox (TOT_UPD & " Entries in SKU - Batch Updated "), vbInformation, "DAS Version 4.0"
            Exit Sub
        End If
        
        If rs.State = 1 Then rs.Close
        rs.Open "select * from PRD where ALT_IDY = '" & ALT_IDY & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            M_PRD_IDY = rs("PRD_IDY")
            rs("PUR_TAX") = Val(PUR_TAX)
            rs("SAL_TAX") = Val(SAL_TAX)
            rs("PUR_DIS") = Val(PUR_DIS)
            rs("SAL_DIS") = Val(SAL_DIS)
            rs("PUR_RAT") = Val(PUR_RAT)
            rs("SAL_RAT") = Val(SAL_RAT)
            rs("MRP_RAT") = Val(MRP_RAT)
            rs.Update
            If BAT_RS.State = 1 Then BAT_RS.Close
            BAT_RS.Open "select * from BRD where PRD_IDY = '" & Trim(M_PRD_IDY) & "'", con, adOpenKeyset, adLockPessimistic
            If BAT_RS.RecordCount > 0 Then
                'Bat_RS("PRD_IDY") = M_PRD_IDY
                'Bat_RS("BAT_IDY") = "BAT" & T7increment("IDY_BAT")
                If UCase(Bat_nme) = "" Then
                    BAT_RS("BAT_NME") = "MRP:" & Trim(str(B_UNT_RAT)) & " " & Format(Date, "YY") & Format(Date, "MM") & Format(Date, "dd")
                Else
                    BAT_RS("BAT_NME") = UCase(Bat_nme)
                End If
                BAT_RS("PRD_PDR") = Val(PUR_RAT)
                BAT_RS("PRD_SDR") = Val(SAL_RAT)
                BAT_RS("PRD_MRP") = Val(MRP_RAT)
                BAT_RS("PRD_QTY") = "0.000"
                BAT_RS("BRD_DOM") = MFG_DAT
                BAT_RS("BRD_DOE") = BBD_DAT
                BAT_RS("B_CRT_RAT") = Val(LST_RAT)
                BAT_RS("B_CRT_DIS") = Val(LST_DIS)
                BAT_RS("P_INC_DIS") = Val(PUR_INC)
                BAT_RS("S_INC_DIS") = Val(SAL_INC)
                BAT_RS("P_PUR_VAT") = 0#
                BAT_RS("P_CRT_RNT") = Val(LST_RAT)
                BAT_RS("P_CRT_VAT") = Val(LST_DIS)
                BAT_RS("P_TOT_VAT") = 0#
                BAT_RS("PUR_LPP") = 0#
                BAT_RS("U_PUR_LPP") = 0#
                BAT_RS("S_SAL_VAT") = 0#
                BAT_RS("S_CRT_RNT") = Val(LST_RAT)
                BAT_RS("S_CRT_VAT") = Val(LST_DIS)
                BAT_RS("S_TOT_VAT") = 0#
                BAT_RS("SAL_LPP") = 0#
                BAT_RS("U_SAL_LPP") = 0#
                BAT_RS("SKU_SBU") = 0#
                BAT_RS("B_PUR_TAX") = 0#
                BAT_RS("B_SAL_TAX") = 0#
                BAT_RS.Update
                TOT_UPD = TOT_UPD + 1
            End If
        End If
        
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_UPD & " Entries in SKU - Batch Updated "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\BAT.TXT") = True Then
        FS_SOU.DeleteFile ("C:\DEX\BAT.TXT")
    End If
End Sub
Private Sub N_BAT_MEX()

Dim rs As New ADODB.Recordset
Dim Pgr_RS As New ADODB.Recordset
Dim BAT_RS As New ADODB.Recordset
Dim ACT_RS As New ADODB.Recordset

Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim TOT_UPD As Integer
Dim STR_VAL As Integer

Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\BAT.TXT") = True Then
    FILE_NAME = "C:\DEX\BAT.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("Master Data Xchange for SKU - Batch Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("Master Data Xchange for SKU - Batch Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim M_CTR As Integer
    Dim FLD_NME As String
    
    Dim M_GRP_IDY As String
    Dim M_PRD_IDY As String
    Dim M_ACT_IDY As String
    
    Dim ACT_NME As String
    Dim GRP_CAT As String
    Dim GRP_CLS As String
    Dim PRD_NME As String
    Dim ALT_IDY As String
    Dim BRD_NME As String
    Dim PRD_PCK As String
    Dim PRD_PKT As String
    Dim PUR_RAT As String
    Dim PUR_TAX As String
    Dim PUR_DIS As String
    Dim SAL_RAT As String
    Dim SAL_TAX As String
    Dim SAL_DIS As String
    Dim MRP_RAT As String
    Dim VOL_WGT As String
    Dim PRD_UOM As String
    Dim LST_RAT As String
    Dim LST_DIS As String
    Dim PUR_INC As String
    Dim SAL_INC As String
    Dim SUB_UNT As Integer
    Dim SAF_LEV As Integer
    Dim SAF_DAY As Integer
    Dim stk_typ As String
    Dim MFG_DAT As Date
    Dim BBD_DAT As Date
    Dim Bat_nme As String
    
    
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\BAT.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        FLD_NME = ""
        M_CTR = 1
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        If Mid(INL_STR, STR_VAL, STR_LEN) = "BAT" Then
            For M_CTR = 1 To 15
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> ";" Then
                        FLD_NME = FLD_NME + Mid(INL_STR, J, 1)
                    Else
                        If M_CTR = 1 Then
                            ALT_IDY = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 2 Then
                            PUR_RAT = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 3 Then
                            SAL_RAT = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 4 Then
                            MRP_RAT = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 5 Then
                            MFG_DAT = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 6 Then
                            BBD_DAT = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 7 Then
                            Bat_nme = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 8 Then
                            PUR_TAX = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 9 Then
                            SAL_TAX = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 10 Then
                            PUR_DIS = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                        If M_CTR = 11 Then
                            SAL_DIS = Trim(FLD_NME)
                            FLD_NME = ""
                        End If
                         If M_CTR = 12 Then
                            LST_RAT = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        If M_CTR = 13 Then
                            LST_DIS = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        If M_CTR = 14 Then
                            PUR_INC = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        If M_CTR = 15 Then
                            SAL_INC = Val(Trim(FLD_NME))
                            FLD_NME = ""
                        End If
                        M_CTR = M_CTR + 1
                    End If
                Next J
            Next M_CTR
        Else
            MsgBox ("Invalid SKU - Batch Record Encountered!, Contact effmcg@consoftsystems.com !"), vbCritical, "DAS Version 4.0"
            MsgBox (TOT_UPD & " Entries in SKU - Batch Updated "), vbInformation, "DAS Version 4.0"
            Exit Sub
        End If
        
        If rs.State = 1 Then rs.Close
        rs.Open "select * from PRD where ALT_IDY = '" & ALT_IDY & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            M_PRD_IDY = rs("PRD_IDY")
            rs("PUR_TAX") = Val(PUR_TAX)
            rs("SAL_TAX") = Val(SAL_TAX)
            rs("PUR_DIS") = Val(PUR_DIS)
            rs("SAL_DIS") = Val(SAL_DIS)
            rs("PUR_RAT") = Val(PUR_RAT)
            rs("SAL_RAT") = Val(SAL_RAT)
            rs("MRP_RAT") = Val(MRP_RAT)
            rs.Update
            If BAT_RS.State = 1 Then BAT_RS.Close
            BAT_RS.Open "select * from BRD", con, adOpenKeyset, adLockPessimistic
            BAT_RS.AddNew
            BAT_RS("PRD_IDY") = M_PRD_IDY
            BAT_RS("BAT_IDY") = "BAT" & T7increment("IDY_BAT")
            BAT_RS("BAT_NME") = UCase(Bat_nme)
            BAT_RS("PRD_PDR") = Val(PUR_RAT)
            BAT_RS("PRD_SDR") = Val(SAL_RAT)
            BAT_RS("PRD_MRP") = Val(MRP_RAT)
            BAT_RS("PRD_QTY") = "0.000"
            BAT_RS("BRD_DOM") = MFG_DAT
            BAT_RS("BRD_DOE") = BBD_DAT
            BAT_RS("B_CRT_RAT") = Val(LST_RAT)
            BAT_RS("B_CRT_DIS") = Val(LST_DIS)
            BAT_RS("P_INC_DIS") = Val(PUR_INC)
            BAT_RS("S_INC_DIS") = Val(SAL_INC)
            BAT_RS("P_PUR_VAT") = 0#
            BAT_RS("P_CRT_RNT") = Val(LST_RAT)
            BAT_RS("P_CRT_VAT") = Val(LST_DIS)
            BAT_RS("P_TOT_VAT") = 0#
            BAT_RS("PUR_LPP") = 0#
            BAT_RS("U_PUR_LPP") = 0#
            BAT_RS("S_SAL_VAT") = 0#
            BAT_RS("S_CRT_RNT") = Val(LST_RAT)
            BAT_RS("S_CRT_VAT") = Val(LST_DIS)
            BAT_RS("S_TOT_VAT") = 0#
            BAT_RS("SAL_LPP") = 0#
            BAT_RS("U_SAL_LPP") = 0#
            BAT_RS("SKU_SBU") = 0#
            BAT_RS("B_PUR_TAX") = 0#
            BAT_RS("B_SAL_TAX") = 0#
            BAT_RS.Update
            TOT_UPD = TOT_UPD + 1
        End If
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_UPD & " Entries in SKU - Batch Updated "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\BAT.TXT") = True Then
        FS_SOU.DeleteFile ("C:\DEX\BAT.TXT")
    End If
End Sub

Private Sub SHL_MEX()

Dim rs As New ADODB.Recordset
Dim Pgr_RS As New ADODB.Recordset
Dim Shl_RS As New ADODB.Recordset
Dim ACT_RS As New ADODB.Recordset

Dim FILE_NAME As String
Dim STR_LEN As Integer
Dim TOT_REC As Integer
Dim TOT_UPD As Integer
Dim STR_VAL As Integer

Dim FS_SOU As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")
If FS_SOU.FileExists("C:\DEX\SHL.TXT") = True Then
    FILE_NAME = "C:\DEX\SHL.TXT"
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox ("Master Data Xchange for Glass - Shell Not Available!"), vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Else
      MsgBox ("Master Data Xchange for Glass - Shell Not Available!"), vbInformation, "DAS Version 4.0"
       Exit Sub
    End If
    Dim M_CTR As Integer
    Dim FLD_NME As String
    
    Dim M_GRP_IDY As String
    Dim M_PRD_IDY As String
    Dim M_ACT_IDY As String
    Dim M_SHL_IDY As String
    
    Dim shl_nme As String
    Dim GRP_CAT As String
    Dim GRP_CLS As String
    Dim PRD_NME As String
    Dim ALT_IDY As String
    Dim BRD_NME As String
    Dim PRD_PCK As String
    Dim PRD_PKT As String
    Dim PUR_RAT As Integer
    Dim PUR_TAX As Integer
    Dim PUR_DIS As Integer
    Dim SAL_RAT As Integer
    Dim SAL_TAX As Integer
    Dim SAL_DIS As Integer
    Dim MRP_RAT As Integer
    Dim VOL_WGT As Integer
    Dim PRD_UOM As String
    Dim SUB_UNT As Integer
    Dim SAF_LEV As Integer
    Dim SAF_DAY As Integer
    Dim stk_typ As String
    Dim MFG_DAT As Date
    Dim BBD_DAT As Date
    Dim Bat_nme As String
    
    
    Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream, TXT_STREAM As TextStream
    Set fil1 = fso.GetFile("C:\DEX\SHL.TXT")
    Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
    For i = 1 To TOT_REC
        J = 0
        FLD_NME = ""
        M_CTR = 1
        INL_STR = TXT_STREAM.ReadLine
        STR_LEN = Len(Trim(INL_STR))
        STR_VAL = (Len(Trim(INL_STR)) - 2)
        If Mid(INL_STR, STR_VAL, STR_LEN) = "SHL" Then
            For M_CTR = 1 To 2
                For J = 1 To STR_LEN
                    If Mid(INL_STR, J, 1) <> ";" Then
                        FLD_NME = FLD_NME + Mid(INL_STR, J, 1)
                    Else
                        If M_CTR = 1 Then
                            ALT_IDY = FLD_NME
                            FLD_NME = ""
                        End If
                        If M_CTR = 2 Then
                            shl_nme = FLD_NME
                            FLD_NME = ""
                        End If
                        M_CTR = M_CTR + 1
                    End If
                Next J
            Next M_CTR
        Else
            MsgBox ("Invalid Glass - Shell Record Encountered!, Contact effmcg@consoftsystems.com !"), vbCritical, "DAS Version 4.0"
            MsgBox (TOT_UPD & " Entries in Glass - Shell Updated "), vbInformation, "DAS Version 4.0"
            Exit Sub
        End If
        If rs.State = 1 Then rs.Close
        rs.Open "select * from PRD where ALT_IDY = '" & ALT_IDY & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            M_PRD_IDY = rs("PRD_IDY")
            M_SHL_IDY = ""
            If Shl_RS.State = 1 Then Shl_RS.Close
            Shl_RS.Open "select * from SHL where SHL_NME = '" & UCase(Trim(shl_nme)) & "'", con, adOpenKeyset, adLockPessimistic
            If Shl_RS.EOF = True Then
                Shl_RS.AddNew
                Shl_RS("PRD_IDY") = rs("PRD_IDY")
                Shl_RS("SHL_IDY") = increment("IDY_SHL")
                Shl_RS("SHL_NME") = UCase(shl_nme)
                Shl_RS("UNT_RAT") = "0.00"
                Shl_RS("SHL_TYP") = "Shell"
                Shl_RS.Update
                TOT_UPD = TOT_UPD + 1
            Else
                M_SHL_IDY = Shl_RS("SHL_IDY")
                Shl_RS.AddNew
                Shl_RS("SHL_IDY") = M_SHL_IDY
                Shl_RS("PRD_IDY") = rs("PRD_IDY")
                Shl_RS("SHL_NME") = UCase(shl_nme)
                Shl_RS("UNT_RAT") = "0.00"
                Shl_RS("SHL_TYP") = "Shell"
                Shl_RS.Update
                TOT_UPD = TOT_UPD + 1
            End If
        End If
        TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    MsgBox (TOT_UPD & " Entries in Glass - Shell Updated "), vbInformation, "DAS Version 4.0"
    If FS_SOU.FileExists("C:\DEX\SHL.TXT") = True Then
        FS_SOU.DeleteFile ("C:\DEX\SHL.TXT")
    End If
End Sub
Private Sub CRT_MEX()
If M_CRT_RNT < 0 Then
    MsgBox ("Invalid Crate Rental Value, Cannot Process This Value"), vbCritical, "DAS Version 4.0"
    Exit Sub
End If
If M_CRT_VAT >= 100 Then
    MsgBox ("Invalid Crate VAT % Value, Cannot Process This Value"), vbCritical, "DAS Version 4.0"
    Exit Sub
End If
Dim Shl_RS As New ADODB.Recordset
Dim PRD_RS As New ADODB.Recordset
If PRD_RS.State = 1 Then PRD_RS.Close
PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
If PRD_RS.RecordCount > 0 Then
    While Not PRD_RS.EOF
        PRD_RS("LST_RAT") = 0#
        PRD_RS("LST_DIS") = 0
        PRD_RS.Update
        PRD_RS.MoveNext
    Wend
End If
If Shl_RS.State = 1 Then Shl_RS.Close
Shl_RS.Open "select * from SHL", con, adOpenKeyset, adLockPessimistic
If Shl_RS.RecordCount > 0 Then
    While Not Shl_RS.EOF
       If PRD_RS.State = 1 Then PRD_RS.Close
       PRD_RS.Open "select * from PRD where PRD_IDY = '" & Shl_RS("PRD_IDY") & "'", con, adOpenKeyset, adLockPessimistic
        If PRD_RS.EOF = False Then
            PRD_RS("LST_RAT") = Val(M_CRT_RNT)
            PRD_RS("LST_DIS") = Val(M_CRT_VAT)
            PRD_RS.Update
        End If
       Shl_RS.MoveNext
    Wend
End If
MsgBox ("Crate Rentals and VAT% on Crate Rental for RGB SKUs Update Successful !"), vbInformation, "DAS Version 4.0"
End Sub
Function MailSend(sSendTo As String, sSubject As String, sText As String, Optional sAttachFile As String) As Boolean
'Dim MAPISession1 As MAPISession
'Dim MAPIMessages1 As MAPIMessages
    On Error GoTo ErrHandler
    With MAPISession1
        If .SessionID = 0 Then
            .DownLoadMail = False
            .LogonUI = True
            .SignOn
            .NewSession = True
            MAPIMessages1.SessionID = .SessionID
        End If
    End With
    With MAPIMessages1
        .Compose
        .RecipAddress = sSendTo
        .AddressResolveUI = True
        .ResolveName
        .MsgSubject = sSubject
        .MsgNoteText = sText
        If Len(sAttachFile) > 0 And Len(Dir$(sAttachFile)) > 0 Then
            .AttachmentPathName = sAttachFile
        Else
            .AttachmentCount = 0
        End If
        .Send False
    End With
    MailSend = True
    Exit Function
ErrHandler:
    Debug.Print err.Description
    MailSend = False
End Function


