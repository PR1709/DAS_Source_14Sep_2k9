VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_CRN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Note "
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "M_FRM_CRN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdXRef 
      Caption         =   "&Van - MER Reference"
      Height          =   375
      Left            =   6960
      TabIndex        =   23
      Top             =   50
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
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
      TabCaption(0)   =   "&Credit Note Voucher"
      TabPicture(0)   =   "M_FRM_CRN.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
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
      Tab(0).Control(9)=   "lblMERNum"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblVanRef"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "date1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "amount"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "rmk"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dbt"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "crd"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "vch"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_CRN.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label49"
      Tab(1).Control(1)=   "list4"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "list1"
      Tab(1).ControlCount=   4
      Begin VB.TextBox vch 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         TabIndex        =   13
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox crd 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         TabIndex        =   12
         Top             =   1700
         Width           =   1215
      End
      Begin VB.TextBox dbt 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         TabIndex        =   11
         Top             =   2220
         Width           =   1215
      End
      Begin VB.TextBox rmk 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         MaxLength       =   70
         TabIndex        =   10
         Top             =   2700
         Width           =   4845
      End
      Begin VB.TextBox amount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         TabIndex        =   9
         Top             =   3180
         Width           =   1215
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5953
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Voucher No."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Credit A/c"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Debit A/c"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74760
         TabIndex        =   2
         Top             =   4320
         Width           =   8295
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
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   855
         End
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
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5953
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
            Text            =   "Account Name"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   330
         Left            =   2280
         TabIndex        =   14
         Top             =   1140
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   72417283
         CurrentDate     =   37233
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
         Left            =   120
         TabIndex        =   25
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   1440
         TabIndex        =   24
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number"
         Height          =   195
         Left            =   1545
         TabIndex        =   22
         Top             =   735
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   1755
         TabIndex        =   21
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Account"
         Height          =   195
         Left            =   1080
         TabIndex        =   20
         Top             =   1770
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit Account"
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   2295
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         Height          =   195
         Left            =   1470
         TabIndex        =   18
         Top             =   2768
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   1560
         TabIndex        =   17
         Top             =   3248
         Width           =   540
      End
      Begin VB.Label Label8 
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
         Left            =   3600
         TabIndex        =   16
         Top             =   1695
         Width           =   3525
      End
      Begin VB.Label Label9 
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
         Left            =   3600
         TabIndex        =   15
         Top             =   2220
         Width           =   3525
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   2055
         TabIndex        =   7
         Top             =   1000
         Width           =   45
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   6
         Top             =   600
         Width           =   8295
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
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
            Picture         =   "M_FRM_CRN.frx":047A
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRN.frx":09BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRN.frx":0D86
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRN.frx":10DA
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRN.frx":1712
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRN.frx":1A66
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRN.frx":1EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRN.frx":21BA
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRN.frx":25C6
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRN.frx":26D2
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRN.frx":2A26
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CRN.frx":2E3A
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Salesmen ID "
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
Attribute VB_Name = "M_FRM_CRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD1 As Boolean, dlr_idy

Private Sub crdlist()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If Check1.Value = 1 Then
    rs.Open "select * from act where ACT_NME LIKE '" & Trim(crd) & "%' and act_idy not like '" & Sincrement(dbt) & "'", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "select * from act where ACT_AD4 LIKE '" & Trim(crd) & "%' and act_idy not like '" & Sincrement(dbt) & "'", con, adOpenKeyset, adLockPessimistic
End If
list4.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy"))
        ls.ListSubItems.Add , , Trim(rs("act_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label49 = "Select Credit Account"
    list4.Visible = True
    Frame1.Visible = True
    list4.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Account Not Found, Define Account...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    crd.SetFocus
End If
End If
End Sub

Private Sub dbtlist()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If Check1.Value = 1 Then
    rs.Open "select * from act where ACT_NME LIKE '" & Trim(dbt) & "%' and act_idy not like '" & Sincrement(crd) & "'", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "select * from act where  ACT_AD4 LIKE '" & Trim(dbt) & "%' and act_idy not like '" & Sincrement(crd) & "'", con, adOpenKeyset, adLockPessimistic
End If
list4.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(rs("act_nme")))
        ls.ListSubItems.Add , , Trim(rs("act_idy"))
        ls.ListSubItems.Add , , Trim(rs("act_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label49 = "Select Debit Account"
    list4.Visible = True
    Frame1.Visible = True
    list4.SetFocus
Else
If SSTab1.Tab = 0 Then
    MsgBox "Account Not Found, Define Account !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    dbt.SetFocus
End If
End If
End Sub

Private Sub amount_Change()
If Len(Trim(amount)) = 9 Then SendKeys "{home}+{end}"
End Sub

Private Sub amount_GotFocus()
SendKeys "{home}+{end}"
End Sub

Private Sub amount_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    amount = convert1(AMT1(amount))
    saverecord
ElseIf K = 38 Then
    amount = convert1(AMT1(amount))
    rmk.SetFocus
End If
End Sub

Private Sub amount_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
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


Private Sub CHECK3_KeyPress(K As Integer)
If K = 13 Then Command1_Click
End Sub

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If Check3.Value = 0 Then
    con.Execute "delete from prt where prt_nme like 'M_FRM_CRN'"
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select * from prt", con, adOpenKeyset, adLockPessimistic
    rs.AddNew
        rs(0) = "M_FRM_CRN"
    rs.Update
End If
Picture1.Visible = False
SSTab1.Enabled = True
Toolbar1.Enabled = True
SSTab1.SetFocus
End Sub

Private Sub Command2_Click()
Picture1.Visible = False
SSTab1.Enabled = True
Toolbar1.Enabled = True
SSTab1.SetFocus
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

Private Sub crd_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("act.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        CANCELRECORD
        Exit Sub
    End If

    Dim rs As New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "Select act_idy,act_nme from act where act_idy like  '" & Sincrement(crd) & "' and act_idy not like '" & Sincrement(dbt) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        crd = Sincrement(rs(0))
        Label8 = Trim(rs(1))
        If K = 38 Then
            DATE1.SetFocus
        ElseIf K = 13 Or K = 40 Then
            dbt.SetFocus
        End If
    Else
        crdlist
    End If
End If
End Sub

Private Sub date1_CloseUp()
DATE1.SetFocus
End Sub

Private Sub date1_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then crd.SetFocus
End Sub

Private Sub dbt_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Debit Account ID or Press <Enter> to Select Debit Account from the List"
End Sub

Private Sub crd_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Credit Account ID or Press <Enter> to Select Credit Account from the List"
End Sub

Private Sub dbt_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Or K = 38 Then
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
    Dim rs As New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "Select act_idy,act_nme from act where act_idy like  '" & Sincrement(dbt) & "' and act_idy not like '" & Sincrement(crd) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        dbt = Sincrement(rs(0))
        Label9 = Trim(rs(1))
    If K = 38 Then
        crd.SetFocus
    ElseIf K = 13 Or K = 40 Then
        rmk.SetFocus
    End If

    Else
        dbtlist
    End If
End If
End Sub

Private Sub dbt_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub crd_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Form_Activate()
SETYEAR M_FRM_CRN
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
If vch.Enabled = True Then
vch.SetFocus
Else
DATE1.SetFocus
End If
Label10 = labdisp("idy_crn")
lblVanRef.Caption = EFF_LIB1.VanRef
lblMERNum.Caption = EFF_LIB1.MerNum
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If list1.Visible = False And list4.Visible = False Then
    If KeyCode = 78 And Shift = 2 Then
    newrecord
    End If
    If KeyCode = 80 And Shift = 2 Then Printrecord
    If KeyCode = 83 And Shift = 2 Then saverecord
    If KeyCode = 88 And Shift = 2 Then Unload Me
    If KeyCode = 81 And Shift = 2 Then CANCELRECORD
    If KeyCode = 112 Then M_HELP
    If KeyCode = 69 And Shift = 2 Then
        If M_FRM_CRN.ActiveControl.name = "vch" Then
            vchlist
        ElseIf M_FRM_CRN.ActiveControl.name = "dbt" Then
            dbtlist
        ElseIf M_FRM_CRN.ActiveControl.name = "crd" Then
            crdlist
        End If
    End If
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_CRN.ActiveControl Is TextBox Then M_FRM_CRN.ActiveControl.Text = FUNKEY(KeyCode)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        vch.SetFocus
    ElseIf list4.Visible = True Then
        list4.Visible = False
        SSTab1.Tab = 0
        If Label49 = "Select Debit Account" Then
            dbt.SetFocus
        Else
            crd.SetFocus
        End If
    Else
        If M_FRM_CRN.ActiveControl.name = "vch" Then
            Unload Me
        Else
            saverecord
        End If
    End If
End If
End Sub

Private Sub Form_Load()
MOD1 = False
DATE1 = Date
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

End Sub

Private Sub List1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub List1_DblClick()
List1_keypress 13
End Sub

Private Sub List1_keypress(K As Integer)
If K = 13 Then
    vch = list1.SelectedItem
    list1.Visible = False
    SSTab1.Tab = 0
    DISPDETAILS
End If
End Sub

Private Sub List4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If list4.SortOrder = lvwAscending Then
list4.SortOrder = lvwDescending
ElseIf list4.SortOrder = lvwDescending Then
list4.SortOrder = lvwAscending
End If
list4.SortKey = ColumnHeader.Index - 1
End Sub
Private Sub list4_DblClick()
list4_KeyPress 13
End Sub

Private Sub list4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SSTab1.Tab = 0
    list4.Visible = False
    If Label49 = "Select Debit Account" Then
        dbt = list4.SelectedItem.ListSubItems.Item(1)
        Label9 = list4.SelectedItem
        rmk.SetFocus
    Else
        crd = list4.SelectedItem.ListSubItems.Item(1)
        crd_KeyDown 13, 0
    End If
End If
End Sub

Private Sub rmk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    amount.SetFocus
ElseIf KeyCode = 38 Then
    dbt.SetFocus
End If
End Sub

Private Sub rmk_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
If vch.Enabled = True Then
    vch.SetFocus
Else
    DATE1.SetFocus
End If
End If
If SSTab1.Tab = 1 Then
    'Toolbar1.Enabled = False
Else
    Toolbar1.Enabled = True
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
Prefix.Caption = "Prefix Credit Note (Accounts)"
Prefix.Show 1
ElseIf Button.KEY = "c" Then
CANCELRECORD
ElseIf Button.KEY = "ed" Then
    If M_FRM_CRN.ActiveControl.name = "vch" Then
        vchlist
    ElseIf M_FRM_CRN.ActiveControl.name = "crd" Then
        crdlist
    ElseIf M_FRM_CRN.ActiveControl.name = "dbt" Then
        dbtlist
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

If check("prt", "prt_nme", "M_FRM_CRN") = False Then
    Check3.Value = 0
Else
    Check3.Value = 1
End If
Check3.SetFocus
Toolbar1.Enabled = False
SSTab1.Enabled = False
End If
End Sub

Public Sub newrecord()
CANCELRECORD
vch.Enabled = False
SSTab1.Tab = 0
DATE1.SetFocus
End Sub

Public Sub CANCELRECORD()
MOD1 = False
clrctr M_FRM_CRN
Label8 = ""
Label9 = ""
amount = "0.00"
SSTab1.Tab = 0
vch.Enabled = True
vch.SetFocus
DATE1 = Date

Label10 = labdisp("idy_crn")
lblMERNum.Caption = ""
lblVanRef.Caption = ""
End Sub

Private Sub vch_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    If Len(Trim(vch)) <> 0 Then
        DISPDETAILS
    Else
        vchlist
    End If
ElseIf K = 27 Then
    Unload Me
End If
End Sub

Private Sub vch_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Public Sub saverecord()
If MsgBox("Confirm Transaction Save ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
    Exit Sub
End If
If CHECKTB("act.dbf,vch.dbf,prt.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim rs As New ADODB.Recordset, RS1 As New ADODB.Recordset, rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Sincrement(crd) = Sincrement(dbt) And crd <> "" Then
    MsgBox "Credit A/c and Debit A/c Cannot Be Same", vbExclamation, "DAS Version 4.0"
    crd = ""
    SSTab1.Tab = 0
    crd.SetFocus
    Exit Sub
End If

If check("act", "act_idy", crd) = False Then
    MsgBox "Account Not Found, Define Account...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    crd.SetFocus
    Exit Sub
End If
If check("act", "act_idy", dbt) = False Then
    MsgBox "Account Not Found, Define Account...!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    dbt.SetFocus
    Exit Sub
End If

If rs.State = 1 Then rs.Close
If MOD1 = True Then
    rs.Open "Select * from vch where txn_idy like '" & Trim(vch) & "' and txn_typ like 'CRN' ", con, adOpenKeyset, adLockPessimistic
    INC = Trim(vch)
Else
    rs.Open "Select * from vch where txn_idy = ''", con, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs("txn_idy") = "CRN" & T7increment("idy_crn")
    INC = rs("txn_idy")
    INVIDY = INC
End If
rs("txn_dat") = DATE1
rs("crt_idy") = Sincrement(crd)
rs("dbt_idy") = Sincrement(dbt)
rs("txn_rmk") = Trim(rmk)
rs("txn_amt") = convert1(AMT1(amount))
rs("txn_typ") = "CRN"
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
If MOD1 = False Then MsgBox "Credit Note Number is " & INC, vbExclamation, "DAS Version 4.0"
If check("PRT", "PRT_NME", "M_FRM_CRN") = True Then
    con.Execute "update dir where dir_idy like 'INV' set dir_als = '" & INC & "'"
    OpenReport CURDIR & "\vouch.RP1", "New CRT NOTE", 2, "vch->txn_idy='" & Trim(INC) & "'.and.vch->txn_typ='CRN'", 6
End If
CANCELRECORD
End Sub

Public Sub DISPDETAILS()
On Error GoTo LBL
Dim rs As New ADODB.Recordset, RS1 As New ADODB.Recordset
If CHECKTB("vch.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If rs.State = 1 Then rs.Close
If Len(Trim(vch)) <> 10 Then
    rs.Open "Select * from vch where txn_idy like 'CRN" & Tincrement(vch) & "' and txn_typ like 'CRN'", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "Select * from vch where txn_idy like '" & Trim(vch) & "' and txn_typ like 'CRN'", con, adOpenKeyset, adLockPessimistic
End If
If rs.RecordCount > 0 Then
    If Len(Trim(rs("VAN_REF"))) > 0 Or Len(Trim(rs("MER_NUM"))) > 0 Then
        MsgBox ("This Voucher has Van Ref " & Trim(rs("VAN_REF")) & " MER Reference " & Trim(rs("MER_NUM"))), vbInformation, "DAS Version 4.0"
    End If
    MOD1 = True
    If Len(Trim(vch)) <> 10 Then vch = "CRN" & Tincrement(vch)
    DATE1 = Trim(rs("txn_dat"))
    dbt = Trim(rs("dbt_idy"))
    If RS1.State = 1 Then RS1.Close
    RS1.Open "Select act_nme from act where act_idy like '" & Sincrement(dbt) & "'", con, adOpenKeyset, adLockPessimistic
    If RS1.RecordCount > 0 Then Label9 = Trim(RS1(0))
    crd = Trim(rs("crt_idy"))
    dlr_idy = Trim(crd)
    If RS1.State = 1 Then RS1.Close
    RS1.Open "Select act_nme from act where act_idy like '" & Sincrement(crd) & "'", con, adOpenKeyset, adLockPessimistic
    If RS1.RecordCount > 0 Then Label8 = Trim(RS1(0))
    rmk = Trim(rs("txn_rmk"))
    amount = convert1(AMT1(rs("txn_amt")))
'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
        If IsNull(rs("VAN_REF")) = False Then lblVanRef.Caption = Trim(rs("VAN_REF"))
        If IsNull(rs("MER_NUM")) = False Then lblMERNum.Caption = Trim(rs("MER_NUM"))
'=====================================================================
'===========================  SYAM  ==================================
'=====================================================================
vch.Enabled = False
    DATE1.SetFocus
Else
    MsgBox "Credit Note Number Not Found !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    vch.Enabled = True
    vch.SetFocus
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

Public Sub vchlist()
If CHECKTB("vch.dbf,act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim VDATE1, VDATE2 As String
If rs.State = 1 Then rs.Close
rs.Open "select vch.txn_idy,vch.txn_dat,vch.txn_amt,a1.act_nme,a2.act_nme from vch,act a1,act a2 where vch.txn_typ like 'CRN' and a1.act_idy like vch.dbt_idy and a2.act_idy like vch.crt_idy", con, adOpenKeyset, adLockPessimistic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs(0)))
        ls.ListSubItems.Add , , Trim(rs(4))
        ls.ListSubItems.Add , , Trim(rs(3))
        VDATE1 = Format(rs(1), "MM/DD/YYYY")
        VDATE2 = Format(VDATE1, "DD/MM/YYYY")
        ls.ListSubItems.Add , , Trim(VDATE2)
        ls.ListSubItems.Add , , convert1(AMT1(Trim(rs(2))))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label49 = "Select Credit Note Voucher"
    list1.Visible = True
    Frame1.Visible = False
    list1.SetFocus
Else
    MsgBox "Credit Note Number Not Found !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    vch.Enabled = True
    vch.SetFocus
End If
End Sub

Public Sub Printrecord()
M_FRM_PRV.Tag = "CRN"
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    CANCELRECORD
    Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_CRN") = True Then M_FRM_PRV.Check3.Value = 1
M_FRM_PRV.Show 1
End Sub
