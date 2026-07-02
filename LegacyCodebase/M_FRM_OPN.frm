VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_OPN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Opening Balance"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   Icon            =   "M_FRM_OPN.frx":0000
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
            Picture         =   "M_FRM_OPN.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPN.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPN.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPN.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPN.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPN.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPN.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPN.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPN.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPN.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPN.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_OPN.frx":2E02
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
      TabCaption(0)   =   "&Opening Balance Details"
      TabPicture(0)   =   "M_FRM_OPN.frx":334E
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
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "date1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "A/c &Postings"
      TabPicture(1)   =   "M_FRM_OPN.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "M_ACT_NME"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "MS"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "amount"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "View"
      TabPicture(2)   =   "M_FRM_OPN.frx":3386
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
         TabIndex        =   18
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
            TabIndex        =   20
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
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   120
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView list2 
         Height          =   3495
         Left            =   -74760
         TabIndex        =   17
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   0
         Top             =   780
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   330
         Left            =   2505
         TabIndex        =   2
         Top             =   1380
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   14737632
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   82640899
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
         TabIndex        =   21
         Top             =   4800
         Width           =   45
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
         TabIndex        =   16
         Top             =   480
         Width           =   8295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3840
         TabIndex        =   15
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -66645
         TabIndex        =   12
         Top             =   4845
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -68400
         TabIndex        =   11
         Top             =   4845
         Width           =   1035
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Voucher Amount"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   660
         TabIndex        =   7
         Top             =   1980
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Number"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1050
         TabIndex        =   6
         Top             =   848
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1905
         TabIndex        =   5
         Top             =   1448
         Width           =   345
      End
      Begin VB.Label label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2205
         TabIndex        =   4
         Top             =   1020
         Width           =   45
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
         Left            =   2520
         TabIndex        =   3
         Top             =   1905
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
   End
End
Attribute VB_Name = "M_FRM_OPN"
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
Dim REC_IRT As New ADODB.Recordset
SSTab1.Tab = 0
If MsgBox("Confirm Transaction Save ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
    Exit Sub
End If
If CHECKTB("vch.dbf,prt.dbf,dir.dbf,irt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
For A = 1 To MS.Rows - 1
    If Len(Trim(MS.TextMatrix(A, 0))) = 0 Then
        MsgBox "Invalid Data, Debit Account Should not be Empty", vbInformation
        SSTab1.Tab = 1
        MS.Col = 0
        MS.SetFocus
        MS.Row = A
        Exit Sub
    ElseIf Len(Trim(MS.TextMatrix(A, 1))) = 0 Then
        MsgBox "Invalid Data, Credit Account Should not be Empty", vbInformation
        SSTab1.Tab = 1
        MS.SetFocus
        MS.Row = A
        MS.Col = 1
        Exit Sub
    End If
Next
Dim RS1 As New ADODB.Recordset
Dim ACTRS As New ADODB.Recordset
Dim IRTRS As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim EDBTIDY As String
If MOD1 = True Then
    con.Execute "delete from vch where txn_idy like '" & (Text1) & "' and txn_typ='OPN'"
    INC = Trim(Text1)
Else
    INC = "OPN" & T7increment("idy_OPN")
End If
If MOD1 = True Then
    con.Execute "delete from IRT where INV_idy like '" & (Text1) & "' and txn_typ='OPN'"
    INC = Trim(Text1)
'Else
'    INC = "OPN" & T7increment("idy_OPN")
End If
If rs.State = 1 Then rs.Close
rs.Open " select * from vch where txn_idy = ''", con, adOpenKeyset, adLockPessimistic
For A = 1 To MS.Rows - 1
If MS.TextMatrix(A, 0) <> "" And Val(Trim(MS.TextMatrix(A, 2))) >= "0.000" Then
rs.AddNew
rs("txn_idy") = INC
rs("txn_dat") = date1
rs("dbt_idy") = Trim(MS.TextMatrix(A, 0))

'CODE ADDED BY PKB
'EDBTIDY = ""
'EDBTIDY = Trim(MS.TextMatrix(A, 0))
'If ACTRS.State = 1 Then ACTRS.Close
'ACTRS.Open " select * from ACT where ACT_idy = '" & EDBTIDY & "' AND TRIM(GRP_CLF) = 'Debtor'", CON, adOpenKeyset, adlockpessimistic
'If ACTRS.RecordCount > 0 Then
'
'If IRTRS.State = 1 Then IRTRS.Close
'IRTRS.Open " select * from IRT WHERE INV_idy LIKE '' ", CON, adOpenKeyset, adlockpessimistic
'IRTRS.AddNew
'IRTRS("INV_idy") = INC
'IRTRS("INV_dat") = date1
'IRTRS("ACT_IDY") = " "
'IRTRS("PTY_idy") = Trim(MS.TextMatrix(A, 0))
'IRTRS("txn_typ") = "OPN"
'IRTRS("NET_amt") = AMT1((MS.TextMatrix(A, 2)))
'IRTRS("DUE_DAT") = date1
'IRTRS("SMN_idy") = " "
'IRTRS.Update
'End If
'END OF CODE SEGMENT

rs("crt_idy") = Trim(MS.TextMatrix(A, 1))
rs("txn_typ") = "OPN"
rs("txn_amt") = AMT1((MS.TextMatrix(A, 2)))
rs("txn_rmk") = Trim(MS.TextMatrix(A, 3))
rs("CLG_DAT") = date1
rs("inv_idy") = ""
rs.Update
End If
Next
If MOD1 = False Then MsgBox "Voucher Number is " & INC, vbExclamation, "DAS Version 4.0"
If check("PRT", "PRT_NME", "M_FRM_OPN") = True Then
con.Execute "update dir where dir_idy like 'INV' set dir_als = '" & INC & "'"
OpenReport CURDIR & "\vouch.RP1", "New OPN Voucher", 2, "vch->txn_idy='" & Trim(INC) & "'.and.vch->txn_typ='OPN'", 6
End If
CANCELRECORD
End Sub

Public Sub CANCELRECORD()
clrctr M_FRM_OPN
SSTab1.Tab = 0
Text1.Enabled = True
Text1.SetFocus
Label23 = "0.00"
Label9 = "0.00"
Label5 = ""
MS.Rows = 1
MS.AddItem ""
label2 = labdisp("idy_OPN")
idnew = False
date1 = Date
MOD1 = False
End Sub

Public Sub newrecord()
CANCELRECORD
Text1.Enabled = False
date1.SetFocus
End Sub

Private Sub Check3_Click()
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub

Private Sub Command1_Click()
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If Check3.Value = 0 Then
    con.Execute "delete from prt where prt_nme like 'M_FRM_OPN'"
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select * from prt", con, adOpenKeyset, adLockPessimistic
    rs.AddNew
        rs(0) = "M_FRM_OPN"
    rs.Update
End If
Picture1.Visible = False
SSTab1.Enabled = True
Toolbar1.Enabled = True
End Sub

Private Sub Command2_Click()
Picture1.Visible = False
End Sub

Private Sub date1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Date using UP or DOWN ARROW KEYS"
End Sub

Private Sub date1_KeyDown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then
    MS.SetFocus
    SSTab1.Tab = 1
End If
End Sub

Private Sub date1_CloseUp()
date1.SetFocus
End Sub

Private Sub Form_Activate()
checkdongle
SSTab1.Tab = 0
SSTab1.TabEnabled(2) = False
If Text1.Enabled = True Then Text1.SetFocus
label2 = labdisp("idy_OPN")
SETYEAR M_FRM_OPN
date1 = Date
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If list1.Visible = False And list2.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then newrecord
    If KeyCode = 80 And Shift = 2 Then Printrecord
    If KeyCode = 83 And Shift = 2 Then saverecord
    If KeyCode = 88 And Shift = 2 Then Unload Me
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
'    If KeyCode = 112 Then Shell "c:\windows\hh.exe " & App.Path & "\conthelp.chm", vbMaximizedFocus
    If KeyCode = 69 And Shift = 2 Then
        If M_FRM_OPN.ActiveControl.name = "Text1" Then
            vchlist
        ElseIf M_FRM_OPN.ActiveControl.name = "Text3" Then
            dbtlist
        ElseIf M_FRM_OPN.ActiveControl.name = "Text3" And MS.Col = 1 Then
            crdlist
        End If
    End If
End If

'If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_OPN.ActiveControl Is TextBox Then CREDITSALES.ActiveControl.Text = FUNKEY(KeyCode)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        Text1.Enabled = True
        Text1.SetFocus
    ElseIf list2.Visible = True Then
        list2.Visible = False
        If label49 = "Select Credit Account" Then
            SSTab1.Tab = 1
            MS.SetFocus
        ElseIf label49 = "Select Debit Account" Then
            SSTab1.Tab = 1
            MS.SetFocus
        End If
    ElseIf SSTab1.Tab = 1 Then
        Text3.Visible = False
        amount.Visible = False
        Label23 = "0.00"
        For A = 1 To MS.Rows - 1
        Label23 = convert1(Label23 + Val(MS.TextMatrix(A, 1)))
        Next
        saverecord
    Else
        If M_FRM_OPN.ActiveControl.name = "Text1" Then
            Unload Me
        Else
            saverecord
        End If
    End If
End If
End Sub

Private Sub Form_Load()
MS.ColWidth(0) = 1200
MS.ColWidth(1) = 1200
MS.ColWidth(2) = 1500
MS.ColWidth(3) = 5000
MS.ColWidth(4) = 2500
MS.ColWidth(5) = 2500
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
    If label49 = "Select Credit Account" Then
        MS.Text = list2.SelectedItem.ListSubItems.Item(1)
        r = MS.Row
        MS.TextMatrix(MS.Row, 5) = list2.SelectedItem
        Text3.Visible = False
        SSTab1.Tab = 1
        MS.SetFocus
        MS.Col = MS.Col + 2
        MS.Row = r
    ElseIf label49 = "Select Debit Account" Then
        MS.Text = list2.SelectedItem.ListSubItems.Item(1)
        r = MS.Row
        MS.TextMatrix(MS.Row, 4) = list2.SelectedItem
        Text3.Visible = False
        SSTab1.Tab = 1
        MS.SetFocus
        MS.Col = MS.Col + 1 'C1
        MS.Row = r
    End If
End If
End Sub

Private Sub list2_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list2.SortOrder = lvwAscending Then
list2.SortOrder = lvwDescending
ElseIf list2.SortOrder = lvwDescending Then
list2.SortOrder = lvwAscending
End If
list2.SortKey = CH.Index - 1
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
ElseIf MS.Col = 3 Then
        Text3.Visible = True
        Text3 = ""
        Text3.Top = MS.Top + MS.CellTop
        Text3.Left = MS.CellLeft + MS.Left
        Text3.Width = MS.CellWidth
        Text3.Height = MS.CellHeight
        Text3.MaxLength = 45
        Text3.Visible = True
        Text3.SetFocus
        If K <> 13 Then Text3 = UCase(Chr(K))
        Text3.SelStart = Len(Text3)
ElseIf MS.Col = 2 Then
    If K > 47 And K < 58 Or K = 46 Or K = 13 Then
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
    Else
        K = 0
    End If
End If
If MS.Col = 3 And MS.Row = MS.Rows - 1 Then
    If MS.TextMatrix(MS.Row, 0) <> "" And K = 13 Then
    list2.Visible = False
    MS.AddItem ""
    MS.Row = MS.Rows - 1
    MS.Col = 0
    End If
End If
End If
If K = 27 And MS.TextMatrix(MS.Row, 0) <> "" Then
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
        date1.SetFocus
    End If
End If
If SSTab1.Tab = 1 Then
    If MS.Rows > 1 Then
        MS.SetFocus
        MS.Row = 1
        MS.Col = 0
    End If
ElseIf SSTab1.Tab = 2 Then
    'Toolbar1.Enabled = False
Else
    Toolbar1.Enabled = True
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
If K = 39 Or K = 41 Then K = 0
If K = 13 Then
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim rs As New ADODB.Recordset
    If MS.Col = 3 Then
        MS.TextMatrix(MS.Row, 3) = Trim(Text3)
        Text3.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 0 Then
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM ACT WHERE ACT_IDY LIKE '" & Sincrement(Text3) & "' and act_idy not like '" & Trim(MS.TextMatrix(MS.Row, 1)) & "'", con, adOpenStatic
        If rs.RecordCount > 0 Then
            r = MS.Row
            MS.TextMatrix(MS.Row, 0) = Sincrement(Text3)
            MS.TextMatrix(MS.Row, 4) = Trim(rs("act_nme"))
            Text3.Visible = False
            MS.SetFocus
            MS.Col = MS.Col + 1
            MS.Row = r
        Else
            dbtlist
        End If
    ElseIf MS.Col = 1 Then
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM ACT WHERE ACT_IDY LIKE '" & Sincrement(Text3) & "' and act_idy not like '" & Trim(MS.TextMatrix(MS.Row, 0)) & "'", con, adOpenStatic
        If rs.RecordCount > 0 Then
            r = MS.Row
            MS.TextMatrix(MS.Row, 1) = Sincrement(Text3)
            MS.TextMatrix(MS.Row, 5) = Trim(rs("act_nme"))
            Text3.Visible = False
            MS.SetFocus
            MS.Col = MS.Col + 1
            MS.Row = r
        Else
            dbtlist
        End If
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
rs.Open "select * from vch where txn_idy like 'OPN" & Tincrement(Text1) & "' and txn_typ like 'OPN'", con, adOpenStatic
Else
rs.Open "select * from vch where txn_idy like '" & Trim(Text1) & "' and txn_typ like 'OPN'", con, adOpenStatic
End If
If rs.RecordCount > 0 Then
    If Len(Trim(Text1)) <> 10 Then Text1 = "OPN" & Tincrement(Text1)
    MOD1 = True
    date1 = Trim(rs("txn_dat"))
    Text1.Enabled = False
    SSTab1.Tab = 0
    date1.SetFocus
    MS.Rows = rs.RecordCount + 1
    For A = 1 To rs.RecordCount
        MS.TextMatrix(A, 0) = Trim(rs("dbt_idy") & "")
        MS.TextMatrix(A, 1) = Trim(rs("crt_idy") & "")
        MS.TextMatrix(A, 2) = convert1(AMT1(rs("txn_amt")))
        MS.TextMatrix(A, 3) = Trim(rs("txn_rmk") & "")
        If RS1.State = 1 Then RS1.Close
        RS1.Open "Select act_nme from act where act_idy like '" & Trim(rs("dbt_idy") & "") & "'", con, adOpenKeyset, adLockPessimistic
        If RS1.RecordCount > 0 Then MS.TextMatrix(A, 4) = Trim(RS1("act_nme"))
        If RS1.State = 1 Then RS1.Close
        RS1.Open "Select act_nme from act where act_idy like '" & Trim(rs("crt_idy") & "") & "'", con, adOpenKeyset, adLockPessimistic
        If RS1.RecordCount > 0 Then MS.TextMatrix(A, 5) = Trim(RS1("act_nme"))
        rs.MoveNext
    Next
    If RS1.State = 1 Then RS1.Close
    RS1.Open "select sum(txn_amt) from vch where txn_idy like '" & (Text1) & "' and txn_typ='OPN'", con, adOpenStatic
    Label23 = convert1(AMT1(RS1(0)))
    Label9 = convert1(AMT1(RS1(0)))
Else
    MsgBox "Voucher Number Not Found !", vbExclamation, "DAS Version 4.0"
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
list2.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list2.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy"))
        ls.ListSubItems.Add , , Trim(rs("act_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    label49 = "Select Credit Account"
    list2.Visible = True
    Frame1.Visible = True
    list2.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
    Text3.Visible = False
    MS.SetFocus
    SSTab1.Tab = 1
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
    rs.Open "select * from act where ACT_NME LIKE '" & Trim(Text3) & "%' and act_idy not like '" & Sincrement(MS.TextMatrix(MS.Row, 1)) & "'", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "select * from act where ACT_AD4 LIKE '" & Trim(Text3) & "%' and act_idy not like '" & Sincrement(MS.TextMatrix(MS.Row, 1)) & "'", con, adOpenKeyset, adLockPessimistic
End If
list2.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list2.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy"))
        ls.ListSubItems.Add , , Trim(rs("act_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    label49 = "Select Debit Account"
    list2.Visible = True
    Frame1.Visible = True
    list2.SetFocus
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
Dim rs As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim VDATE1, VDATE2 As String
If rs.State = 1 Then rs.Close
rs.Open "select (vch.txn_idy),vch.txn_dat,act.act_nme from vch,act where vch.txn_typ like 'OPN' and act.act_idy like vch.crt_idy GROUP BY TXN_IDY", con, adOpenKeyset, adLockPessimistic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs(0)))
        If RS1.State = 1 Then RS1.Close
        RS1.Open "Select sum(txn_amt) from vch where txn_idy like '" & Sincrement(rs(0)) & "' and txn_typ like 'OPN'", con, adOpenKeyset, adLockPessimistic
        ls.ListSubItems.Add , , Trim(rs(2))
        VDATE1 = Format(rs(1), "MM/DD/YYYY")
        VDATE2 = Format(VDATE1, "DD/MM/YYYY")
        ls.ListSubItems.Add , , Trim(VDATE2)
        'ls.ListSubItems.Add , , Trim(rs(1))
        ls.ListSubItems.Add , , convert1(AMT1(Trim(RS1(0))))
        rs.MoveNext
    Next
    SSTab1.Tab = 2
    label49 = "Select A/c Opening Balance Voucher"
    list1.Visible = True
    Frame1.Visible = False
    list1.SetFocus
Else
    MsgBox "Voucher Number Not Found !", vbExclamation, "DAS Version 4.0"
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
If label49.Caption = "Select Debit Account" Then
    dbtlist
ElseIf label49.Caption = "Select Credit Account" Then
    crdlist
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
If label49.Caption = "Select Debit Account" Then
    dbtlist
ElseIf label49.Caption = "Select Credit Account" Then
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
Prefix.Caption = "Prefix Opening Balance"
Prefix.Show 1
ElseIf Button.KEY = "c" Then
CANCELRECORD
ElseIf Button.KEY = "ed" Then
    If M_FRM_OPN.ActiveControl.name = "Text3" And MS.Col = 1 Then
        crdlist
    ElseIf M_FRM_OPN.ActiveControl.name = "Text3" Then
        dbtlist
    ElseIf M_FRM_OPN.ActiveControl.name = "Text1" Then
        vchlist
    End If
ElseIf Button.KEY = "h" Then
'Shell "c:\windows\hh.exe " & App.Path & "\conthelp.chm", vbMaximizedFocus
ElseIf Button.KEY = "pre" Then
Picture1.Visible = True
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("prt", "prt_nme", "M_FRM_OPN") = False Then
    Check3.Value = 0
Else
    Check3.Value = 1
End If
Check3.SetFocus
'Toolbar1.Enabled = False
SSTab1.Enabled = False
End If
End Sub

Public Sub Printrecord()
M_FRM_PRV.Tag = "OPN"
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_OPN") = True Then M_FRM_PRV.Check3.Value = 1
M_FRM_PRV.Show 1
End Sub
