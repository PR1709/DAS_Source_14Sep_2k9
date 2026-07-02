VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_JRN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Journal Entries"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   Icon            =   "M_FRM_JRN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
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
            Picture         =   "M_FRM_JRN.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_JRN.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_JRN.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_JRN.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_JRN.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_JRN.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_JRN.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_JRN.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_JRN.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_JRN.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_JRN.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_JRN.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5145
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      TabCaption(0)   =   "&Journal Voucher"
      TabPicture(0)   =   "M_FRM_JRN.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label23"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label22"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblVanRef"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblMERNum"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "date1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "&Detail Entries"
      TabPicture(1)   =   "M_FRM_JRN.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(2)=   "M_ACT_NME"
      Tab(1).Control(3)=   "MS"
      Tab(1).Control(4)=   "Text3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "amount"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "View"
      TabPicture(2)   =   "M_FRM_JRN.frx":3386
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "label49"
      Tab(2).Control(1)=   "list1"
      Tab(2).Control(2)=   "list2"
      Tab(2).Control(3)=   "Frame1"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74760
         TabIndex        =   17
         Top             =   4440
         Width           =   8295
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Place"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3480
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   120
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   120
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView list2 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   16
         Top             =   840
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
            Text            =   "Account ID"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   4057
         EndProperty
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2625
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1200
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   330
         Left            =   2625
         TabIndex        =   2
         Top             =   1800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   14737632
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74186755
         CurrentDate     =   37099
         MaxDate         =   37256
         MinDate         =   36526
      End
      Begin MSMask.MaskEdBox amount 
         Height          =   375
         Left            =   -69000
         TabIndex        =   9
         Top             =   540
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         MaxLength       =   9
         Format          =   "##,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   -68040
         MaxLength       =   10
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1140
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   4380
         Left            =   -74955
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   7726
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   2
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   13
         Top             =   840
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Voucher "
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Credit A/C"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label lblMERNum 
         Alignment       =   1  'Right Justify
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
         Left            =   1560
         TabIndex        =   22
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblVanRef 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         TabIndex        =   21
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label M_ACT_NME 
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
         Left            =   -74760
         TabIndex        =   20
         Top             =   4845
         Width           =   1125
      End
      Begin VB.Label label49 
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
         TabIndex        =   15
         Top             =   480
         Width           =   8295
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Height          =   195
         Left            =   -67080
         TabIndex        =   12
         Top             =   4845
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount :"
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
         Height          =   195
         Left            =   -68760
         TabIndex        =   11
         Top             =   4845
         Width           =   1260
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   7
         Top             =   2475
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   6
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2115
         TabIndex        =   5
         Top             =   1875
         Width           =   345
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   3960
         TabIndex        =   4
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
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
         Left            =   2625
         TabIndex        =   3
         Top             =   2400
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   14
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
            Object.Visible         =   0   'False
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Account ID "
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
      Begin VB.CommandButton cmdXRef 
         Caption         =   "&Van - MER Reference"
         Height          =   375
         Left            =   7040
         TabIndex        =   23
         Top             =   0
         Width           =   1815
      End
   End
End
Attribute VB_Name = "M_FRM_JRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idnew As Boolean, MOD1 As Boolean, ls As ListItem

Private Sub amount_Change()
If Len(Trim(amount)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub amount_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 8 Then
ElseIf K = 13 Then
    MS.Text = convert1(AMT1(Val(amount)))
    amount.Visible = False
    MS.SetFocus
    MS.Col = MS.Col + 1
Else
K = 0
End If
End Sub

Public Sub saverecord()
Dim rs As New ADODB.Recordset, RS1 As New ADODB.Recordset, rs2 As New ADODB.Recordset
SSTab1.Tab = 0
If MsgBox("Confirm Transaction Save ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
    Exit Sub
End If
If CHECKTB("vch.dbf,prt.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
For A = 1 To MS.Rows - 1
    If Len(Trim(MS.TextMatrix(A, 0))) = 0 Then
        MsgBox "Debit A/c ID Required!", vbInformation
        SSTab1.Tab = 1
        MS.Col = 0
        MS.SetFocus
        MS.Row = A
        Exit Sub
    ElseIf Len(Trim(MS.TextMatrix(A, 1))) = 0 Then
        MsgBox "Credit A/c ID Required!", vbInformation
        SSTab1.Tab = 1
        MS.SetFocus
        MS.Row = A
        MS.Col = 1
        Exit Sub
    End If
Next
If MOD1 = True Then
    con.Execute "delete from vch where txn_idy like '" & Trim(Text1) & "' and txn_typ='JRN'"
    INC = Trim(Text1)
Else
    INC = "JRN" & T7increment("IDY_JRN")
End If
If rs.State = 1 Then rs.Close
rs.Open " select * from vch where txn_idy = ''", con, adOpenKeyset, adLockPessimistic
For A = 1 To MS.Rows - 1
If MS.TextMatrix(A, 0) <> "" And Val(Trim(MS.TextMatrix(A, 2))) >= "0.000" Then
rs.AddNew
rs("txn_idy") = INC
rs("txn_dat") = DATE1
rs("dbt_idy") = Trim(MS.TextMatrix(A, 0))
rs("crt_idy") = Trim(MS.TextMatrix(A, 1))
rs("txn_typ") = "JRN"
rs("txn_amt") = convert1(AMT1(MS.TextMatrix(A, 2)))
rs("txn_rmk") = Trim(MS.TextMatrix(A, 3))
rs("CLG_DAT") = DATE1
rs("inv_idy") = ""
'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
                rs("VAN_REF") = Trim(lblVanRef.Caption)
                rs("MER_NUM") = Trim(lblMERNum.Caption)

'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
rs.Update
End If
Next
If MOD1 = False Then MsgBox "Voucher Number is " & INC, vbExclamation, "DAS Version 4.0"
If check("PRT", "PRT_NME", "M_FRM_JRN") = True Then
con.Execute "update dir where dir_idy like 'INV' set dir_als = '" & INC & "'"
OpenReport CURDIR & "\vouch.RP1", "New JRN Voucher", 2, "vch->txn_idy='" & Trim(INC) & "'.and.vch->txn_typ='JRN'", 6
End If
CANCELRECORD
End Sub

Public Sub CANCELRECORD()
clrctr M_FRM_JRN
SSTab1.Tab = 0
Text1.Enabled = True
Text1.SetFocus
Label23 = "0.00"
Label9 = "0.00"
Label5 = ""
MS.Rows = 1
MS.AddItem ""
Label2 = labdisp("IDY_JRN")
idnew = False
DATE1 = Date
MOD1 = False
lblMERNum.Caption = ""
lblVanRef.Caption = ""
End Sub

Public Sub newrecord()
CANCELRECORD
Text1.Enabled = False
DATE1.SetFocus
End Sub

Private Sub cmdXRef_Click()
'If Text1.Text = "" Then
'Text1.SetFocus
'Exit Sub
'End If
MerNum = lblMERNum.Caption
VanRef = lblVanRef.Caption
M_FRM_XRF.Show 1
End Sub

Private Sub date1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Date using UP or DOWN ARROW KEYS"
End Sub

Private Sub date1_KeyDown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then
    SSTab1.Tab = 1
    MS.SetFocus
    MS.Col = 0
End If
End Sub

Private Sub date1_CloseUp()
DATE1.SetFocus
End Sub

Private Sub Form_Activate()
checkdongle
SSTab1.Tab = 0
SSTab1.TabEnabled(2) = False
If Text1.Enabled = True Then Text1.SetFocus
Label2 = labdisp("idy_JRN")
SETYEAR M_FRM_JRN
DATE1 = Date
lblVanRef.Caption = EFF_LIB1.VanRef
lblMERNum.Caption = EFF_LIB1.MerNum
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If
If list1.Visible = False And List2.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then newrecord
    If KeyCode = 80 And Shift = 2 Then Printrecord
    If KeyCode = 83 And Shift = 2 Then saverecord
    If KeyCode = 88 And Shift = 2 Then Unload Me
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
        If M_FRM_JRN.ActiveControl.name = "Text1" Then
            vchlist
        ElseIf M_FRM_JRN.ActiveControl.name = "Text3" Then
            If MS.Col = 0 Then dbtlist
            If MS.Col = 1 Then crdlist
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_JRN.ActiveControl Is TextBox Then M_FRM_JRN.ActiveControl.Text = FUNKEY(KeyCode)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        Text1.Enabled = True
        Text1.SetFocus
    ElseIf List2.Visible = True Then
        Text3.Visible = False
        List2.Visible = False
        SSTab1.Tab = 1
        MS.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        Text3.Visible = False
        amount.Visible = False
        Label23 = "0.00"
        For A = 1 To MS.Rows - 1
        Label23 = convert1(Label23 + Val(MS.TextMatrix(A, 2)))
        Next
        saverecord
    Else
        If M_FRM_JRN.ActiveControl.name = "Text1" Then
            Unload Me
        Else
            saverecord
        End If
    End If
End If
End Sub

Private Sub Form_Load()
MS.ColWidth(0) = 1300
MS.ColWidth(1) = 1100
MS.ColWidth(2) = 1100
MS.ColWidth(3) = 6000 'changed by PNR from 5000 to 6000 as per prince chopra request
MS.ColWidth(4) = 2000
MS.ColWidth(5) = 2000

MS.TextMatrix(0, 0) = "Debit A/C"
MS.ColAlignment(0) = 1
SCl MS, 0
MS.TextMatrix(0, 1) = "Credit A/C"
MS.ColAlignment(1) = 1
SCl MS, 1
MS.TextMatrix(0, 2) = "   Amount"
SCl MS, 2
MS.TextMatrix(0, 3) = "Remarks"
SCl MS, 3
MS.TextMatrix(0, 4) = "Debit Account Name"
MS.TextMatrix(0, 5) = "Credit Account Name"
idnew = False
MOD1 = False
Label23 = "0.00"


End Sub

Private Sub Form_Unload(cancel As Integer)
'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
EFF_LIB1.VanRef = ""
EFF_LIB1.MerNum = ""
'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub List1_keypress(K As Integer)
If K = 13 Then
   Text1 = list1.SelectedItem
   list1.Visible = False
   DISPDETAILS
End If
End Sub

Private Sub list1_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub

Private Sub list2_DblClick()
list2_keypress 13
End Sub

Private Sub list2_keypress(K As Integer)
If K = 13 Then
Dim rs As New ADODB.Recordset
    If Label49 = "Select Credit Account" Then
        MS.Text = List2.SelectedItem.ListSubItems.Item(1)
        MS.TextMatrix(MS.Row, 5) = Trim(List2.SelectedItem)
        r = MS.Row
        Text3.Visible = False
        SSTab1.Tab = 1
        MS.SetFocus
        MS.Col = MS.Col + 2
        MS.Row = r
    ElseIf Label49 = "Select Debit Account" Then
        MS.Text = List2.SelectedItem.ListSubItems.Item(1)
        r = MS.Row
        MS.TextMatrix(MS.Row, 4) = Trim(List2.SelectedItem)
        Text3.Visible = False
        SSTab1.Tab = 1
        MS.SetFocus
        MS.Col = MS.Col + 1
        MS.Row = r
    End If
End If
End Sub

Private Sub list2_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If List2.SortOrder = lvwAscending Then
List2.SortOrder = lvwDescending
ElseIf List2.SortOrder = lvwDescending Then
List2.SortOrder = lvwAscending
End If
List2.SortKey = CH.Index - 1
End Sub

Private Sub List1_DblClick()
List1_keypress 13
End Sub

Private Sub MS_EnterCell()
If MS.Col = 3 Then
    Label9 = "0.00"
    For A = 1 To MS.Rows - 1
        Label9 = convert1(Val(Label9) + Val(MS.TextMatrix(A, 2)))
    Next
    M_MDI.STS_BAR.Panels(1).Text = "Enter Remarks"
ElseIf MS.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Debit A/C"
    M_ACT_NME = Trim(MS.TextMatrix(MS.Row, 4))
ElseIf MS.Col = 1 Then
    M_ACT_NME = Trim(MS.TextMatrix(MS.Row, 5))
    M_MDI.STS_BAR.Panels(1).Text = "Enter or Select Credit A/C"
ElseIf MS.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Amount"
ElseIf MS.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Debit A/c Name"
ElseIf MS.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Credit A/c Name"
Else
    M_ACT_NME = ""
End If
End Sub

Private Sub MS_GotFocus()
'MS.HighLight = flexHighlightAlways
End Sub

Private Sub MS_LeaveCell()
Text3.Visible = False
amount.Visible = False
If MS.Visible = True And MS.Enabled = True Then MS.SetFocus
End Sub

Private Sub MS_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
If K <> 27 Then
If MS.Col = 0 Then
    If K = 13 And MS.Text = "" Then
        Text3 = ""
        dbtlist
    ElseIf K = 13 And MS.Text <> "" Then
        MS.Col = MS.Col + 1
    Else
        If K = vbKeySpace Then
        Else
        If Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
        Else
        Text3.Visible = True
        Text3.MaxLength = 10
        Text3.Top = MS.Top + MS.CellTop
        Text3.Left = MS.CellLeft + MS.Left
        Text3.Width = MS.CellWidth
        Text3.Height = MS.CellHeight
        Text3.Visible = True
        Text3.SetFocus
        If K <> 13 Then Text3 = Chr(K)
        Text3.SelStart = Len(Text3)
        End If
        End If
    End If
ElseIf MS.Col = 1 Then
    If K = 13 And MS.Text = "" Then
        Text3 = ""
        crdlist
    ElseIf K = 13 And MS.Text <> "" Then
        MS.Col = MS.Col + 1
    Else
        If K = vbKeySpace Then
        Else
        If Trim(MS.TextMatrix(MS.Row, 1)) <> "" Then
        Else
        Text3.Visible = True
        Text3.MaxLength = 10
        Text3.Top = MS.Top + MS.CellTop
        Text3.Left = MS.CellLeft + MS.Left
        Text3.Width = MS.CellWidth
        Text3.Height = MS.CellHeight
        Text3.Visible = True
        Text3.SetFocus
        If K <> 13 Then Text3 = Chr(K)
        Text3.SelStart = Len(Text3)
        End If
        End If
    End If
ElseIf MS.Col = 2 Then
        If K <> 13 Then amount = Chr(K)
        amount.SelStart = Len(amount)
        amount.Visible = True
        amount = ""
        amount.Text = MS.TextMatrix(MS.Row, 2)
        amount.Top = MS.Top + MS.CellTop
        amount.Left = MS.CellLeft + MS.Left
        amount.Width = MS.CellWidth
        amount.Height = MS.CellHeight
        amount.Visible = True
        amount.SetFocus
        If K <> 13 Then amount = Chr(K)
        amount.SelStart = Len(amount)
ElseIf MS.Col = 3 Then
    If K <> 13 Then
        Text3.Visible = True
        Text3 = ""
        Text3.Top = MS.Top + MS.CellTop
        Text3.Left = MS.CellLeft + MS.Left
        Text3.Width = MS.CellWidth
        Text3.Height = MS.CellHeight
        If K <> 13 Then Text3 = Chr(K)
        Text3.MaxLength = 45
        Text3.Visible = True
        Text3.SetFocus
        Text3.SelStart = Len(Text3)
    End If
End If
If (MS.Col = 3 Or MS.Col = 4 Or MS.Col = 5) And MS.Row = MS.Rows - 1 Then
    If MS.TextMatrix(MS.Row, 0) <> "" And K = 13 Then
    List2.Visible = False
    MS.AddItem ""
    MS.Row = MS.Rows - 1
    MS.Col = 0
    End If
End If
End If
If SSTab1.Tab = 1 And K = 27 And MS.TextMatrix(MS.Row, 0) <> "" Then
saverecord
End If
End Sub
Private Sub MS_Scroll()
If amount.Visible = True Then
    amount = ""
    amount.Visible = False
    MS.SetFocus
ElseIf Text3.Visible = True Then
    Text3 = ""
    Text3.Visible = False
    MS.SetFocus
End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
     If Text1.Enabled = True Then
        Text1.SetFocus
    Else
        DATE1.SetFocus
    End If
End If
If SSTab1.Tab = 1 Then
    If MS.Rows > 1 Then
        MS.SetFocus
        MS.Row = 1
        MS.Col = 0
    End If
ElseIf SSTab1.Tab = 2 Then
   ' Toolbar1.Enabled = False
    Toolbar1.Enabled = True
'    Toolbar1.Enabled = True
End If
End Sub

Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Voucher Number"
End Sub

Private Sub TEXT1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Text1)) <> 0 Then
        DISPDETAILS
    Else
        vchlist
    End If
ElseIf KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub text3_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Then K = 0
If K = 13 Then
    If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If

    Dim rs As New ADODB.Recordset
    If MS.Col = 2 Then
        r = MS.Row
        MS.TextMatrix(MS.Row, 2) = Trim(Text3)
        Text3.Visible = False
        MS.SetFocus
        MS.Row = r
    ElseIf MS.Col = 0 Then
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM ACT WHERE ACT_IDY LIKE '" & Sincrement(Text3) & "' and act_idy not like '" & Trim(MS.TextMatrix(MS.Row, 1)) & "'", con, adOpenStatic
        If rs.RecordCount > 0 Then
            r = MS.Row
            MS.TextMatrix(MS.Row, 0) = Sincrement(Text3)
            MS.TextMatrix(MS.Row, 4) = Trim(rs("act_nme") & "")
            Text3.Visible = False
            MS.SetFocus
            MS.Col = MS.Col + 1
            MS.Row = r
        Else
            dbtlist
        End If
    ElseIf MS.Col = 1 Then
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM ACT WHERE ACT_IDY LIKE '" & Sincrement(Text3) & "' and act_idy not like '" & Trim(Trim(MS.TextMatrix(MS.Row, 0))) & "'", con, adOpenStatic
        If rs.RecordCount > 0 Then
            r = MS.Row
            MS.TextMatrix(MS.Row, 1) = Sincrement(Text3)
            MS.TextMatrix(MS.Row, 5) = Trim(rs("act_nme") & "")
            Text3.Visible = False
            MS.SetFocus
            MS.Col = MS.Col + 1
            MS.Row = r
        Else
            crdlist
        End If
    ElseIf MS.Col = 3 Then
        MS.Text = Text3
        Text3.Visible = False
        MS.SetFocus
    End If
End If
End Sub

Public Sub DISPDETAILS()
On Error GoTo LBL
If CHECKTB("vch.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim rs As New ADODB.Recordset, RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Len(Trim(Text1)) <> 10 Then
rs.Open "select * from vch where txn_idy like 'JRN" & Tincrement(Text1) & "' and txn_typ like 'JRN'", con, adOpenStatic
Else
rs.Open "select * from vch where txn_idy like '" & Trim(Text1) & "' and txn_typ like 'JRN'", con, adOpenStatic
End If
If rs.RecordCount > 0 Then
    If Len(Trim(rs("VAN_REF"))) > 0 Or Len(Trim(rs("MER_NUM"))) > 0 Then
        MsgBox ("This Voucher has Van Ref " & Trim(rs("VAN_REF")) & " MER Reference " & Trim(rs("MER_NUM"))), vbInformation, "DAS Version 4.0"
    End If
    If Len(Trim(Text1)) <> 10 Then Text1 = "JRN" & Tincrement(Text1)
    MOD1 = True
    DATE1 = Trim(rs("txn_dat"))
    Text1.Enabled = False
    Text1 = Sincrement(Text1)
    SSTab1.Tab = 0
    DATE1.SetFocus
    MS.Rows = rs.RecordCount + 1
    For A = 1 To rs.RecordCount
         MS.TextMatrix(A, 0) = Trim(rs("dbt_idy") & "")
         MS.TextMatrix(A, 1) = Trim(rs("CRT_idy") & "")
         MS.TextMatrix(A, 2) = convert1(AMT1(rs("txn_amt")))
         MS.TextMatrix(A, 3) = Trim(rs("txn_rmk") & "")
         If RS1.State = 1 Then RS1.Close
         RS1.Open "Select act_nme from act where act_idy like '" & Sincrement(rs("dbt_idy")) & "'", con, adOpenKeyset, adLockPessimistic
         If RS1.RecordCount > 0 Then MS.TextMatrix(A, 4) = Trim(RS1("act_nme"))
         If RS1.State = 1 Then RS1.Close
         RS1.Open "Select act_nme from act where act_idy like '" & Sincrement(rs("CRt_idy")) & "'", con, adOpenKeyset, adLockPessimistic
         If RS1.RecordCount > 0 Then MS.TextMatrix(A, 5) = Trim(RS1("act_nme"))
'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
        If IsNull(rs("VAN_REF")) = False Then lblVanRef.Caption = Trim(rs("VAN_REF"))
        If IsNull(rs("MER_NUM")) = False Then lblMERNum.Caption = Trim(rs("MER_NUM"))
'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
         rs.MoveNext
    Next
    If RS1.State = 1 Then RS1.Close
    RS1.Open "select sum(txn_amt) from vch where txn_idy like '" & Trim(Text1) & "' and txn_typ='JRN'", con, adOpenStatic
    Label23 = convert1(AMT1(RS1(0)))
    Label9 = convert1(AMT1(RS1(0)))
Else
    MsgBox "Journal Voucher Number Not Found !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text1.Enabled = True
    Text1.SetFocus
End If
Exit Sub
LBL:
If err.Number = 35773 Then
    MsgBox "Invalid Voucher Selected for this Financial Year", vbExclamation
    CANCELRECORD
Else
    MsgBox err.Description
    CANCELRECORD
End If
End Sub

Private Sub crdlist()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Check1.Value = 1 Then
    rs.Open "select * from act where ACT_NME LIKE '" & Trim(Text3) & "%' AND ACT_IDY NOT LIKE '" & Trim(MS.TextMatrix(MS.Row, 0)) & "'", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "select * from act where ACT_AD4 LIKE '" & Trim(Text3) & "%' AND ACT_IDY NOT LIKE '" & Trim(MS.TextMatrix(MS.Row, 0)) & "'", con, adOpenKeyset, adLockPessimistic
End If
List2.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = List2.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy"))
        ls.ListSubItems.Add , , Trim(rs("act_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    Label49 = "Select Credit Account"
    List2.Visible = True
    Frame1.Visible = True
    List2.SetFocus
Else
If SSTab1.Tab = 1 Then
    MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
    Text3.Visible = False
    SSTab1.Tab = 1
    MS.SetFocus
End If
End If
End Sub

Private Sub dbtlist()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Check1.Value = 1 Then
    rs.Open "select * from act where ACT_NME LIKE '" & Trim(Text3) & "%' and act_idy not like '" & Sincrement(Trim(MS.TextMatrix(MS.Row, 1))) & "'", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "select * from act where ACT_AD4 LIKE '" & Trim(Text3) & "%' and act_idy not like '" & Sincrement(Trim(MS.TextMatrix(MS.Row, 1))) & "'", con, adOpenKeyset, adLockPessimistic
End If
List2.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = List2.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy"))
        ls.ListSubItems.Add , , Trim(rs("act_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    Label49 = "Select Debit Account"
    List2.Visible = True
    Frame1.Visible = True
    List2.SetFocus
Else
If SSTab1.Tab = 1 Then
    MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 1
    Text3.Visible = False
    MS.SetFocus
End If
End If
End Sub

Public Sub vchlist()
If CHECKTB("vch.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select vch.txn_idy,vch.txn_dat,sum(vch.txn_amt),act.act_nme from vch,act where vch.txn_typ like 'JRN' and act.act_idy like vch.crt_idy GROUP BY TXN_IDY", con, adOpenKeyset, adLockPessimistic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs(0)))
        ls.ListSubItems.Add , , Trim(rs(3))
        'ls.ListSubItems.Add , , Trim(rs(1))
        VDATE1 = Format(Trim(rs(1)), "MM/DD/YYYY")
        VDATE2 = Format(VDATE1, "DD/MM/YYYY")
        ls.ListSubItems.Add , , VDATE2
       ls.ListSubItems.Add , , convert1(AMT1(Trim(rs(2))))
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    Label49 = "Select Journal Voucher"
    list1.Visible = True
    Frame1.Visible = False
    
    list1.SetFocus
Else
    MsgBox "Journal Voucher Not Found !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
ElseIf Check1.Value = 0 Then
    Check2.Value = 1
End If
If Label49.Caption = "Select Debit Account" Then
    dbtlist
ElseIf Label49.Caption = "Select Credit Account" Then
    crdlist
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
If Label49.Caption = "Select Debit Account" Then
    dbtlist
ElseIf Label49.Caption = "Select Credit Account" Then
    crdlist
End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
newrecord
ElseIf Button.KEY = "s" Then
saverecord
ElseIf Button.KEY = "e" Then
Unload Me
ElseIf Button.KEY = "p" Then
Printrecord
ElseIf Button.KEY = "pre" Then
Prefix.Caption = "Prefix Journal Entries"
Prefix.Show 1
ElseIf Button.KEY = "c" Then
CANCELRECORD
ElseIf Button.KEY = "ed" Then
    If M_FRM_JRN.ActiveControl.name = "Text3" Then
        If MS.Col = 0 Then dbtlist
        If MS.Col = 1 Then dbtlist
    ElseIf M_FRM_JRN.ActiveControl.name = "Text1" Then
        vchlist
    End If
ElseIf Button.KEY = "h" Then
M_HELP
ElseIf Button.KEY = "pre" Then
Picture1.Visible = True
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If

If check("prt", "prt_nme", "M_FRM_JRN") = False Then
    Check3.Value = 0
Else
    Check3.Value = 1
End If
Check3.SetFocus
'Toolbar1.Enabled = False
'SSTab1.Enabled = False
End If
End Sub

Public Sub Printrecord()
M_FRM_PRV.Tag = "JRN"
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_JRN") = True Then M_FRM_PRV.Check3.Value = 1
M_FRM_PRV.Show 1
End Sub
