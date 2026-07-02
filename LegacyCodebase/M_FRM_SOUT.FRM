VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_SOUT 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shells Outward"
   ClientHeight    =   6420
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
   Icon            =   "M_FRM_SOUT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6420
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
            Picture         =   "M_FRM_SOUT.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOUT.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOUT.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOUT.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOUT.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOUT.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOUT.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOUT.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOUT.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOUT.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOUT.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SOUT.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   10610
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
      TabCaption(0)   =   "&Outward Details"
      TabPicture(0)   =   "M_FRM_SOUT.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MS"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CMD_GEN"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_SOUT.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "list1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   8655
         Begin VB.CheckBox PTY_CHK 
            Caption         =   "&Party"
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
            Left            =   2160
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   278
            Width           =   855
         End
         Begin VB.CheckBox SMN_CHK 
            Caption         =   "&S.Man"
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
            Left            =   120
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   278
            Width           =   855
         End
         Begin VB.TextBox SMN_IDY 
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
            Left            =   970
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   210
            Width           =   1095
         End
         Begin VB.TextBox PTY_IDY 
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
            Left            =   3015
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   210
            Width           =   1095
         End
         Begin VB.CheckBox LDS_CHK 
            Caption         =   "&Load No"
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
            Left            =   4200
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   278
            Width           =   975
         End
         Begin VB.TextBox LDS_IDY 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   5175
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   210
            Width           =   1095
         End
         Begin VB.TextBox VAN_IDY 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   7455
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   210
            Width           =   1095
         End
         Begin VB.CheckBox VAN_CHK 
            Caption         =   "&Van Issue"
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
            Left            =   6360
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   278
            Width           =   1095
         End
      End
      Begin VB.CommandButton CMD_GEN 
         Caption         =   "&Generate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3840
         TabIndex        =   10
         Top             =   400
         Width           =   2415
         Begin MSComCtl2.DTPicker FRM_DAT 
            Height          =   330
            Left            =   930
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   315
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24576003
            CurrentDate     =   37258
         End
         Begin MSComCtl2.DTPicker to_dat 
            Height          =   330
            Left            =   930
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24576003
            CurrentDate     =   37258
         End
         Begin VB.Label Label2 
            Caption         =   "To"
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
            Left            =   450
            TabIndex        =   12
            Top             =   795
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "From"
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
            Left            =   330
            TabIndex        =   11
            Top             =   345
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select Transaction"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   400
         Width           =   3495
         Begin VB.OptionButton opt_tsn 
            Caption         =   "&Van Stock Issue"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton opt_tsn 
            Caption         =   "&Debit Notes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton opt_tsn 
            Caption         =   "&Transfer Outward"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton opt_tsn 
            Caption         =   "&Sales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView list1 
         Height          =   4575
         Left            =   -74640
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   8070
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
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   3405
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2520
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   6006
         _Version        =   393216
         Rows            =   1
         Cols            =   14
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   1
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
         TabIndex        =   2
         Top             =   600
         Width           =   8055
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
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
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s"
            Object.ToolTipText     =   "Save (Ctrl+S)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "c"
            Object.ToolTipText     =   "Cancel (Ctrl+Q)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ed"
            Object.ToolTipText     =   "View (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Salesmen ID "
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "brw"
            Object.ToolTipText     =   "Browse (Ctrl+B)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "d"
            Object.ToolTipText     =   "Delete (Ctrl+D)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_SOUT"
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
   If MsgBox("Delete Salesman ?", vbYesNo + vbCritical, "Effmcg") = vbYes Then
        If check("smn", "smn_idy", Trim(Text1)) = False Then
            MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "Effmcg"
            SSTab1.Tab = 0
            clrctr M_FRM_SMN
            Text1.Enabled = True
            Text1.SetFocus
            Exit Sub
        Else
            If deletecheck = False Then
                CON.Execute "delete from smn where smn_idy like '" & Sincrement(Trim(Text1)) & "'"
                SSTab1.Tab = 0
                Text1.Enabled = True
                Text1.SetFocus
                clrctr M_FRM_SMN
            Else
                MsgBox "Cannot Delete, Transaction Exists !", vbExclamation, "Effmcg"
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
Dim RS As New ADODB.Recordset

If MsgBox("Confirm Record Save ?", vbExclamation + vbYesNo, "Effmcg") = vbNo Then
    cancel
    Exit Sub
End If

If Len(Trim(Text2)) = 0 Then
    MsgBox "Blank Name, Enter Name !", vbExclamation, "Effmcg"
    Text2.SetFocus
    Exit Sub
End If
If RS.State = 1 Then RS.Close
If MOD1 = True Then
RS.Open "select * from smn where smn_idy like '" & Sincrement(Trim(Text1)) & "'", CON, adOpenKeyset, adLockOptimistic
Else
RS.Open "select * from smn", CON, adOpenKeyset, adLockOptimistic
RS.AddNew
RS("smn_idy") = increment("idy_smn")
MsgBox "New Salesmen ID is " & RS("smn_idy"), vbExclamation, "Effmcg"
End If
RS("smn_nme") = Trim(Text2)
RS("smn_ad1") = Trim(text3)
RS("smn_ad2") = Trim(Text4)
RS("smn_ad3") = Trim(Text5)
RS("smn_ad4") = Trim(Text6)
RS("smn_tfx") = Trim(Text8)
RS("smn_fax") = Trim(Text9) & ""
RS("smn_mbl") = Trim(Text10)
RS("smn_eml") = Trim(Text11)
RS("smn_zip") = Val(Text7)
RS.Update
If RS.State = 1 Then RS.Close
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
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from smn where smn_nme like '" & Text1 & "%'", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("smn_nme")))
        ls.ListSubItems.Add , , Trim(RS("smn_idy") & "")
        ls.ListSubItems.Add , , Trim(RS("smn_ad4") & "")
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
    Toolbar1.Enabled = False
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        Text1.Enabled = True
        Text1.SetFocus
    End If
End If
End Sub

Private Sub exitfrm()
Unload Me
End Sub

Private Sub Check2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Place To Get List of Salesmen Sorted on Place"
End Sub

Private Sub Check1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Name To Get List of Salesmen Sorted on Name"
End Sub

Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
If opt_tsn(0).Enabled = True Then
    opt_tsn(0).SetFocus
'Else
 '   Text2.SetFocus
End If
End Sub

Private Sub Form_KeyDown(K As Integer, S As Integer)
If K = 120 Then Shell "c:\windows\calc.exe"
If K > 111 And K < 120 Then If TypeOf M_FRM_SMN.ActiveControl Is TextBox Then M_FRM_SMN.ActiveControl.Text = Trim(FUNKEY(K))
If list1.Visible = False Then
    'If K = 78 And S = 2 Then
     '   newrec
    'Else
    If K = 83 And S = 2 Then
        save
    ElseIf K = 81 And S = 2 Then
        cancel
    ElseIf K = 88 And S = 2 Then
        exitfrm
    ElseIf K = 67 And S = 2 Then
        If Toolbar1.Buttons(12).Enabled = True Then CREATE
    ElseIf K = 68 And S = 2 Then
        delete
    ElseIf K = 69 And S = 2 Then
        If M_FRM_SMN.ActiveControl.name = "Text1" Then smnlist
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And S = 2 Then
        Openreport CURDIR & "\lists.rp1", "Salesmen list", 0, "", 6
    ElseIf K = 82 And S = 2 Then
        'MDIForm1.RRReport1.ReportName = App.Path & "\Salesmen list.rrw"
        'MDIForm1.RRReport1.Destination = 1
        'MDIForm1.RRReport1.Printer = "?"
        'MDIForm1.RRReport1.RunReport 1
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    'If list1.Visible = True Then
     '   list1.Visible = False
        SSTab1.Tab = 0
      '  If check("smn", "smn_idy", Trim(Text1)) = False Then
       '     Text1.SetFocus
        'Else
         '   Text1.Enabled = False
          '  Text2.SetFocus
        'End If
        Toolbar1.Enabled = True
    'Else
     '   If M_FRM_SMN.ActiveControl.name = "Text1" Then
           Unload Me
       ' ElseIf M_FRM_SMN.ActiveControl.name = "SSTab1" Then
        'Else
         '   save
        'End If
    'End If
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

Private Sub FRM_DAT_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
to_dat.SetFocus
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
Dim RS As New ADODB.Recordset
If KeyAscii = 13 Then
    If list1.ListItems.Count > 0 Then
        If RS.State = 1 Then RS.Close
        Dim lst As ListSubItem
        Set lst = list1.SelectedItem.ListSubItems.Item(1)
        RS.Open "select * from smn where smn_idy like '" & Trim(lst.Text) & "'", CON, adOpenStatic
        Text1 = Trim(RS("smn_idy"))
        Text2 = Trim(RS("smn_nme"))
        text3 = Trim(RS("smn_ad1") & "")
        Text4 = Trim(RS("smn_ad2") & "")
        Text5 = Trim(RS("smn_ad3") & "")
        Text6 = Trim(RS("smn_ad4") & "")
        Text8 = Trim(RS("smn_tfx") & "")
        Text9 = Trim(RS("smn_fax") & "")
        Text10 = Trim(RS("smn_mbl") & "")
        Text11 = Trim(RS("smn_eml") & "")
        Text7 = Trim(RS("smn_zip") & "")
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
Private Sub opt_tsn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If opt_tsn(Index).Caption = "&Sales" Then
        SMN_CHK.Enabled = True
        SMN_IDY.Enabled = True
        PTY_CHK.Caption = "Party"
        PTY_CHK.Enabled = True
        PTY_IDY.Enabled = True
        LDS_CHK.Enabled = True
        LDS_IDY.Enabled = True
        VAN_CHK.Enabled = True
        VAN_IDY.Enabled = True
        FRM_DAT.SetFocus
    ElseIf opt_tsn(Index).Caption = "&Van Stock Issue" Then
        SMN_CHK.Enabled = True
        SMN_IDY.Enabled = True
        PTY_CHK.Caption = "&Vehicle"
        PTY_CHK.Enabled = True
        'PTY_IDY.Enabled = False
        LDS_CHK.Enabled = False
        LDS_IDY.Enabled = False
        VAN_CHK.Enabled = False
        VAN_IDY.Enabled = False
        FRM_DAT.SetFocus
    Else
        SMN_CHK.Enabled = False
        SMN_IDY.Enabled = False
        'PTY_CHK.Caption = "Party"
        PTY_CHK.Enabled = False
        PTY_IDY.Enabled = False
        LDS_CHK.Enabled = False
        LDS_IDY.Enabled = False
        VAN_CHK.Enabled = False
        VAN_IDY.Enabled = False
        FRM_DAT.SetFocus
    End If
End If
End Sub

Private Sub SMN_CHK_Click()
If SMN_CHK.Value = 1 Then
SMN_IDY.SetFocus
End If
End Sub

Private Sub Smn_IDY_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If KeyCode = 13 Then
    If Trim(SMN_IDY) = "" Then
        RS.Open "select * from smn", CON, adOpenKeyset, adLockOptimistic
        SSTab1.Tab = 1
        With list1
            .View = lvwReport
            .Visible = True
            .SetFocus
            .ColumnHeaders.Add , , "Name"
            .ColumnHeaders.Add , , "ID"
            If RS.RecordCount > 0 Then
                While Not RS.EOF
                    Set ls = .ListItems.Add(, , RS("smn_nme"))
                    ls.ListSubItems.Add , , RS("smn_idy")
                    RS.MoveNext
                Wend
            End If
        End With
    Else
        RS.Open "select * from smn where smn_IDY like '" & Sincrement(Trim(SMN_IDY.Text)) & "'", CON, adOpenStatic
        If RS.RecordCount > 0 Then
            SMN_IDY = Sincrement(Trim(SMN_IDY))
'        DisplayDetails
'        PTY_CHK.Value = 0
'        PTY_IDY = ""
'        VAN_CHK.Value = 0
'        VAN_IDY = ""
'        LDS_CHK.Value = 0
'        LDS_IDY = ""
'    Else
'        V_SMN_LST
        End If
    End If
End If
End Sub

'Private Sub SSTab1_Click(PreviousTab As Integer)
'Toolbar1.Enabled = True
'list1.Visible = False
'If SSTab1.Tab = 0 Then
'    If Text1.Enabled = True Then
'        Text1.SetFocus
'    Else
'        Text2.SetFocus
'    End If
'End If
'End Sub

Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Salesmen ID or Press <Enter> To Get The List of Salesmen ID's"
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
If K = 27 And Text1 = "" Then Unload Me
End Sub

Private Sub Text10_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Mobile or Pager Number"
End Sub

Private Sub Text11_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Email ID"
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

Private Sub Text1_Keydown(K As Integer, Shift As Integer)
Dim RS As New ADODB.Recordset
If K = 13 Or K = 40 Then
    If RS.State = 1 Then RS.Close
    RS.Open "select * from smn where smn_idy like '" & Sincrement(Trim(Text1)) & "'", CON, adOpenStatic
    If RS.RecordCount > 0 Then
        Text1 = Sincrement(Trim(Text1))
        Text2 = Trim(RS("smn_nme"))
        text3 = Trim(RS("smn_ad1") & "")
        Text4 = Trim(RS("smn_ad2") & "")
        Text5 = Trim(RS("smn_ad3") & "")
        Text6 = Trim(RS("smn_ad4") & "")
        Text7 = Trim(RS("smn_zip") & "")
        Text9 = Trim(RS("smn_fax") & "")
        Text10 = Trim(RS("smn_mbl") & "")
        Text11 = Trim(RS("smn_eml") & "")
        Text8 = Trim(RS("smn_tfx") & "")
        MOD1 = True
        Text1.Enabled = False
        Text2.SetFocus
        Toolbar1.Buttons(12).Enabled = True
    Else
        smnlist
    End If
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Text2)) > 0 Then
        text3.SetFocus
    Else
        MsgBox "Blank Name, Enter Name !", vbExclamation, "Effmcg"
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
    text3.SetFocus
End If
End Sub

Private Sub Text5_keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    Text6.SetFocus
ElseIf KeyCode = 38 Then
    Text4.SetFocus
End If
End Sub

Private Sub Text6_keydown(KeyCode As Integer, Shift As Integer)
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
If KeyCode = 13 Or KeyCode = 40 Then
    save
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

Private Sub to_dat_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
SMN_CHK.SetFocus
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'If Button.KEY = "n" Then
 '   newrec
'Else
If Button.KEY = "s" Then
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
    M_FRM_PRE.Caption = "Salesmen ID Prefix"
    M_FRM_PRE.Show 1
    M_MDI.STS_BAR.Panels(1).Text = "Set Prefix for Salesmen ID"
ElseIf Button.KEY = "p" Then
    Openreport CURDIR & "\lists.rp1", "Salesmen list", 0, "", 6
ElseIf Button.KEY = "pr" Then
ElseIf Button.KEY = "crt" Then
    CREATE
End If

End Sub
