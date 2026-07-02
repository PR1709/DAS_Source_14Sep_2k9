VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_TRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Returnables Accounting (Empties Returnables & Actuals)"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "M_FRM_TRS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5145
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
      TabCaption(0)   =   "&Transaction Details"
      TabPicture(0)   =   "M_FRM_TRS.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "MS"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtbrk"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmbTrans"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "invid"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "m_flx_amt"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_TRS.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lv"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.TextBox m_flx_amt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   11
         Top             =   4080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox invid 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         TabIndex        =   9
         Top             =   1030
         Width           =   1815
      End
      Begin VB.ComboBox cmbTrans 
         Height          =   315
         ItemData        =   "M_FRM_TRS.frx":047A
         Left            =   1680
         List            =   "M_FRM_TRS.frx":0490
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtbrk 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComctlLib.ListView lv 
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
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   2835
         Left            =   60
         TabIndex        =   3
         Top             =   2280
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   5001
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   2
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
      Begin VB.Label Label7 
         Caption         =   "Label3"
         Height          =   855
         Left            =   7680
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3600
         TabIndex        =   14
         Top             =   1755
         Width           =   2805
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3600
         TabIndex        =   13
         Top             =   1395
         Width           =   2805
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3600
         TabIndex        =   12
         Top             =   1035
         Width           =   2805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select ID"
         Height          =   195
         Left            =   840
         TabIndex        =   8
         Top             =   1095
         Width           =   660
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Invoice"
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
         TabIndex        =   5
         Top             =   720
         Width           =   7900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Type"
         Height          =   195
         Left            =   255
         TabIndex        =   4
         Top             =   660
         Width           =   1245
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   240
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
            Picture         =   "M_FRM_TRS.frx":04E2
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TRS.frx":0A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TRS.frx":0DEE
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TRS.frx":1142
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TRS.frx":177A
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TRS.frx":1ACE
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TRS.frx":1F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TRS.frx":2222
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TRS.frx":262E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TRS.frx":273A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TRS.frx":2A8E
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_TRS.frx":2EA2
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   8
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ACT_LST"
                  Text            =   "&1 All Accounts List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_CAT"
                  Text            =   "&2 Retailer Category"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_LOC"
                  Text            =   "&3 Retailer Location"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_MAR"
                  Text            =   "&4 Retailer Market"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_ROU"
                  Text            =   "&5 Retailer Beat"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_SMN"
                  Text            =   "&6 Retailer Salesman"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SUP_LST"
                  Text            =   "&7 Supplier List"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_ID_LST"
                  Text            =   "&8 Retailer ID List"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "act"
                  Text            =   "&Accounts List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dlr"
                  Text            =   "&Dealers List"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sup"
                  Text            =   "&Supplier List"
               EndProperty
            EndProperty
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
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Account ID "
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Begin VB.CommandButton Command1 
         Caption         =   "&Sort Empties"
         Height          =   400
         Left            =   6840
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   1935
      End
   End
End
Attribute VB_Name = "M_FRM_TRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ttype As String
Private Sub cmbTrans_Click()
Select Case cmbTrans.ListIndex
    Case 0: ttype = "PUR"
    Case 1: ttype = "SAL"
    Case 2: ttype = "SRN"
    Case 3: ttype = "PRT"
    Case 4: ttype = "STI"
    Case 5: ttype = "STO"
End Select
End Sub
Private Sub cmbTrans_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    invid.SetFocus
End If
End Sub
Private Sub Command1_Click()
EMP_SHL_SOR "TRS_SHL"
End Sub
Private Sub Form_Activate()
If cmbTrans.Enabled = True Then cmbTrans.SetFocus
SSTab1.TabEnabled(1) = False
If Len(Gltxninvid) > 0 Then
    loaddetails
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If SSTab1.Tab = 1 Then
        If lv.Visible = True Then
            lv.Visible = False
            SSTab1.Tab = 0
            If invid.Enabled = True Then
                invid.SetFocus
                SLT
            End If
        End If
    ElseIf Len(Trim(invid)) = 0 Then
        Unload Me
    Else
        saverecord
    End If
End If
End Sub
Private Sub Form_Load()
cmbTrans.ListIndex = 0
MS.Cols = 24
MS.FixedCols = 3
MS.MergeCells = flexMergeFree
MS.Rows = 1
MS.ColWidth(0) = 0
MS.ColWidth(1) = 1500
MS.ColWidth(2) = 1800
For A = 3 To 15
    MS.ColWidth(A) = 1200
Next
MS.MergeCol(0) = True
MS.MergeCol(1) = True
MS.TextMatrix(0, 0) = "Voucher No"
MS.TextMatrix(0, 1) = "Shell Name"
MS.TextMatrix(0, 2) = "SKU Pack - UPC"
MS.ColAlignment(2) = 1
MS.TextMatrix(0, 3) = "Shells"
MS.TextMatrix(0, 4) = "Cases"
MS.TextMatrix(0, 5) = "Recd.Shl"
MS.TextMatrix(0, 6) = "Recd.Bot"
MS.TextMatrix(0, 7) = "Shls Act"
SCl MS, 7
MS.TextMatrix(0, 8) = "Cas Act"
SCl MS, 8
MS.TextMatrix(0, 9) = "Shls Brk"
SCl MS, 9
MS.TextMatrix(0, 10) = "Cas Brk"
SCl MS, 10
MS.TextMatrix(0, 11) = "Shells +/-"
SCl MS, 11
MS.TextMatrix(0, 12) = "Cases +/-"
SCl MS, 12
MS.TextMatrix(0, 13) = "Shls Dep"
SCl MS, 13
MS.TextMatrix(0, 14) = "Cas Dep"
SCl MS, 14

MS.TextMatrix(0, 15) = "Amount"
SCl MS, 15
MS.ColWidth(16) = 0
MS.ColWidth(17) = 0
MS.ColWidth(18) = 0
MS.ColWidth(19) = 0
MS.ColWidth(20) = 0
MS.ColWidth(21) = 0
MS.ColWidth(22) = 0
MS.ColWidth(23) = 0
MS.ColWidth(5) = 0
MS.ColWidth(6) = 0
For A = 3 To 15
    MS.ColAlignment(A) = 7
Next
MS.TextMatrix(0, 16) = "prdidy"
MS.TextMatrix(0, 17) = "shlidy"
End Sub
Private Sub invid_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
'If CHECKTB("snv.dbf") = 1 Then
'    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
''    cancel
'    Exit Sub
'End If
If Len(Trim(invid)) = 0 Then
    V_INV_LST
Else
' LOAD VALUES TO THE FLEXI GRID USING THE INVID
'DISPLIST
End If
End If
End Sub
Private Sub invid_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 47 And K < 58 Or K = 8 Or K = 13 Then
ElseIf K > 96 And K < 123 Then
K = K - 32
ElseIf K > 64 And K < 91 Then
Else
K = 0
End If
If invid = "" And K = 27 Then Unload Me
End Sub
Private Sub V_INV_LST()
If CHECKTB("act.dbf,snv.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'        cancelrecord
        Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select snv.inv_idy,snv.inv_dat,act.act_idy,act.act_nme,act.act_ad1,act.act_ad2,act.act_ad3,act.act_ad4 from act,snv where snv.cst_idy=act.act_idy and snv.txn_typ ='" & Trim(ttype) & "' group by inv_idy", con, adOpenKeyset, adLockPessimistic
lv.ColumnHeaders.CLEAR
lv.ListItems.CLEAR
lv.ColumnHeaders.Add , , "Invoice ID", 1500
lv.ColumnHeaders.Add , , "Date", 1200
lv.ColumnHeaders.Add , , "Retailer-Supplier Name", 3000
lv.ColumnHeaders.Add , , "Place", 1500
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = lv.ListItems.Add(, , Trim(REC_SET(0)))
        ls.ListSubItems.Add , , Format(REC_SET(1), "dd/MM/yyyy")
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_AD4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    lv.Visible = True
    lv.SetFocus
Else
    MsgBox "Invoice Number Not Found!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    invid.Enabled = True
    SendKeys "{HOME}+{END}"
    invid.SetFocus
End If
End Sub
Private Sub LV_DblClick()
lv_KeyPress 13
End Sub
Private Sub lv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
displaydetails Trim(lv.SelectedItem), ttype
End If
End Sub
Private Sub M_FLX_AMT_KeyPress(K As Integer)
Dim sunt As Integer
sunt = Val(Val(MS.TextMatrix(MS.Row, 22)))
If K = 13 Then
    If MS.Col = 9 Or MS.Col = 10 Then
        If MS.Col = 9 Then
'            If (Val(m_flx_amt) + Val(MS.TextMatrix(MS.Row, 7)) + Val(MS.TextMatrix(MS.Row, 11))) <= Val(MS.TextMatrix(MS.Row, 3)) Then
                MS.Text = Int(Val(M_FLX_AMT))
'            Else
                'MsgBox "Invalid Value", vbExclamation + vbOKOnly, "DAS Version 4.0"
                'm_flx_amt.SetFocus
                'SendKeys "{home}+{end}"
                'Exit Sub
            'End If
        ElseIf MS.Col = 10 Then
            If CSUBUNT(Val(M_FLX_AMT), Val(MS.TextMatrix(MS.Row, 22))) = True Then
                'If Val(MTS(Val(m_flx_amt), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 8)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 12)), Val(sunt))) <= MTS(Val(MS.TextMatrix(MS.Row, 4)), Val(sunt)) Then
                    MS.Text = QTY(Val(M_FLX_AMT))
                    M_FLX_AMT.Visible = False
'                Else
'                    MsgBox "Invalid Returns", vbExclamation + vbOKCancel, "DAS Version 4.0"
'                    m_flx_amt.SetFocus
'                    SendKeys "{home}+{end}"
'                    Exit Sub
'                End If
                'MS.Text = QTY(Val(M_FLX_AMT))
                'M_FLX_AMT.Visible = False
            Else
                Exit Sub
            End If
        End If
        MS.TextMatrix(MS.Row, 11) = Val(MS.TextMatrix(MS.Row, 3)) - (Val(MS.TextMatrix(MS.Row, 5)) + Val(MS.TextMatrix(MS.Row, 7)) + Val(MS.TextMatrix(MS.Row, 9)))
        MS.TextMatrix(MS.Row, 12) = QTY(STM(Val(MTS(Val(MS.TextMatrix(MS.Row, 4)), Val(sunt)) - (MTS(Val(MS.TextMatrix(MS.Row, 6)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 8)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 10)), Val(sunt)))), Val(sunt)))
        MS.TextMatrix(MS.Row, 13) = 0
        MS.TextMatrix(MS.Row, 14) = 0
        MS.TextMatrix(MS.Row, 15) = 0
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 7 Or MS.Col = 8 Then
        If MS.Col = 7 Then
            'If (Val(m_flx_amt) + Val(MS.TextMatrix(MS.Row, 9)) + Val(MS.TextMatrix(MS.Row, 11))) <= Val(MS.TextMatrix(MS.Row, 3)) Then
                MS.Text = Int(Val(M_FLX_AMT))
                'MS.TextMatrix(MS.Row, 8) = QTY(Trim(MS.TextMatrix(MS.Row, 7))) 'MS.TextMatrix(MS.Row, 8) = MTS(Val(MS.Text), GetValue("SUB_UNT", "PRD WHERE PRD_IDY LIKE '" & Trim(MS.TextMatrix(MS.Row, 16)) & "'"))
            'Else
             '   MsgBox "Invalid Value", vbExclamation + vbOKOnly, "DAS Version 4.0"
             '   m_flx_amt.SetFocus
             '   SendKeys "{home}+{end}"
             '   Exit Sub
            'End If
        ElseIf MS.Col = 8 Then
            If CSUBUNT(Val(M_FLX_AMT), Val(MS.TextMatrix(MS.Row, 22))) = True Then
                'If Val(MTS(Val(m_flx_amt), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 10)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 12)), Val(sunt))) <= MTS(Val(MS.TextMatrix(MS.Row, 4)), Val(sunt)) Then
                    MS.Text = QTY(Val(M_FLX_AMT))
                    M_FLX_AMT.Visible = False
                'Else
                 '   MsgBox "Invalid Returns", vbExclamation + vbOKCancel, "DAS Version 4.0"
                 '   m_flx_amt.SetFocus
                 '   SendKeys "{home}+{end}"
                 '   Exit Sub
                'End If
            Else
                Exit Sub
            End If
        End If
        MS.TextMatrix(MS.Row, 11) = Val(MS.TextMatrix(MS.Row, 3)) - (Val(MS.TextMatrix(MS.Row, 5)) + Val(MS.TextMatrix(MS.Row, 7)) + Val(MS.TextMatrix(MS.Row, 9)))
        MS.TextMatrix(MS.Row, 12) = QTY(STM(Val(MTS(Val(MS.TextMatrix(MS.Row, 4)), Val(sunt)) - (MTS(Val(MS.TextMatrix(MS.Row, 6)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 8)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 10)), Val(sunt)))), Val(sunt)))
        MS.TextMatrix(MS.Row, 13) = 0
        MS.TextMatrix(MS.Row, 14) = 0
        MS.TextMatrix(MS.Row, 15) = 0
        If cmbTrans.ListIndex = 3 Or cmbTrans.ListIndex = 4 Or cmbTrans.ListIndex = 9 Or cmbTrans.ListIndex = 10 Or cmbTrans.ListIndex = 11 Then
            'If Trim(MS.TextMatrix(MS.Rows - 1, 1)) <> "" Then MS.AddItem ""
            If MS.Col = 8 Then
                If MS.Row + 1 < MS.Rows Then MS.Row = MS.Row + 1
                MS.Col = 7
            ElseIf MS.Col = 7 Then
                MS.Col = 8
            End If
        End If
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 11 Or MS.Col = 12 Then
        If MS.Col = 11 Then
            'If (Val(m_flx_amt) + Val(MS.TextMatrix(MS.Row, 9)) + Val(MS.TextMatrix(MS.Row, 7))) <= Val(MS.TextMatrix(MS.Row, 3)) Then
                MS.Text = Int(Val(M_FLX_AMT))
            'Else
            '    MsgBox "Invalid Value", vbExclamation + vbOKOnly, "DAS Version 4.0"
            '    m_flx_amt.SetFocus
            '    SendKeys "{home}+{end}"
             '   Exit Sub
            'End If
        ElseIf MS.Col = 12 Then
            If CSUBUNT(Val(M_FLX_AMT), Val(MS.TextMatrix(MS.Row, 22))) = True Then
              '  If Val(MTS(Val(m_flx_amt), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 10)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 8)), Val(sunt))) <= MTS(Val(MS.TextMatrix(MS.Row, 4)), Val(sunt)) Then
                    
            
              '  Else
              '      MsgBox "Invalid Returns", vbExclamation + vbOKCancel, "DAS Version 4.0"
              '      m_flx_amt.SetFocus
              '      SendKeys "{home}+{end}"
               '     Exit Sub
               ' End If
                'MS.Text = QTY(Val(M_FLX_AMT))
                'M_FLX_AMT.Visible = False
            Else
                Exit Sub
            End If
        End If
        MSTEXT = MS.Text
        'MS.Text = QTY(Val(M_FLX_AMT))
        MS13 = Val(MS.TextMatrix(MS.Row, 3)) - (Val(MS.TextMatrix(MS.Row, 5)) + Val(MS.TextMatrix(MS.Row, 7)) + Val(MS.TextMatrix(MS.Row, 9)) + Val(MS.TextMatrix(MS.Row, 11)))
        MS14 = QTY(STM((MTS(Val(MS.TextMatrix(MS.Row, 4)), Val(MS.TextMatrix(MS.Row, 22)))) - (MTS(Val(MS.TextMatrix(MS.Row, 6)), Val(MS.TextMatrix(MS.Row, 22))) + MTS(Val(MS.TextMatrix(MS.Row, 8)), Val(MS.TextMatrix(MS.Row, 22))) + MTS(Val(MS.TextMatrix(MS.Row, 10)), Val(MS.TextMatrix(MS.Row, 22))) + MTS(Val(MS.TextMatrix(MS.Row, 12)), Val(MS.TextMatrix(MS.Row, 22)))), Val(MS.TextMatrix(MS.Row, 22))))
        If MS13 < 0 Or MS14 < 0 Then
            MS.Text = MSTEXT
            MsgBox "Invalid Crate Deposit !", vbExclamation, "DAS Version 4.0"
            M_FLX_AMT.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        Else
            M_FLX_AMT.Visible = False
            If MS.Col = 11 Then
                MS.Text = Int(M_FLX_AMT)
            Else
                MS.Text = QTY(Val(M_FLX_AMT))
            End If
            MS.TextMatrix(MS.Row, 13) = Val(MS.TextMatrix(MS.Row, 3)) - (Val(MS.TextMatrix(MS.Row, 5)) + Val(MS.TextMatrix(MS.Row, 7)) + Val(MS.TextMatrix(MS.Row, 9)) + Val(MS.TextMatrix(MS.Row, 11)))
            MS.TextMatrix(MS.Row, 14) = QTY(STM((MTS(Val(MS.TextMatrix(MS.Row, 4)), Val(MS.TextMatrix(MS.Row, 22)))) - (MTS(Val(MS.TextMatrix(MS.Row, 6)), Val(MS.TextMatrix(MS.Row, 22))) + MTS(Val(MS.TextMatrix(MS.Row, 8)), Val(MS.TextMatrix(MS.Row, 22))) + MTS(Val(MS.TextMatrix(MS.Row, 10)), Val(MS.TextMatrix(MS.Row, 22))) + MTS(Val(MS.TextMatrix(MS.Row, 12)), Val(MS.TextMatrix(MS.Row, 22)))), Val(MS.TextMatrix(MS.Row, 22))))
        End If
        M_FLX_AMT.Visible = False
        MS.SetFocus
    ElseIf MS.Col = 3 Or MS.Col = 4 Then
        'If cmbTrans.ListIndex = 0 Or cmbTrans.ListIndex = 1 Or cmbTrans.ListIndex = 3 Or cmbTrans.ListIndex = 4 Or cmbTrans.ListIndex = 9 Or cmbTrans.ListIndex = 10 Or cmbTrans.ListIndex = 11 Then
            'If Trim(MS.TextMatrix(MS.Rows - 1, 1)) <> "" Then MS.AddItem ""
            If MS.Col = 4 Then
                sunt = Trim(MS.TextMatrix(MS.Row, 22))
                If CSUBUNT(Val(M_FLX_AMT), Val(sunt)) = True Then
                    MS.Text = QTY(Val(M_FLX_AMT))
                    MS.TextMatrix(MS.Row, 11) = Val(MS.TextMatrix(MS.Row, 3)) - (Val(MS.TextMatrix(MS.Row, 5)) + Val(MS.TextMatrix(MS.Row, 7)) + Val(MS.TextMatrix(MS.Row, 9)))
                    MS.TextMatrix(MS.Row, 12) = QTY(STM(Val(MTS(Val(MS.TextMatrix(MS.Row, 4)), Val(sunt)) - (MTS(Val(MS.TextMatrix(MS.Row, 6)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 8)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 10)), Val(sunt)))), Val(sunt)))
                    If MS.Row + 1 < MS.Rows Then MS.Row = MS.Row + 1
                    MS.Col = 6
                    M_FLX_AMT.Visible = False
                    MS.SetFocus
                End If
            ElseIf MS.Col = 3 Then
'                MS.TextMatrix(MS.Row, 4) = MTS(Val(MS.Text), Val(MS.TextMatrix(MS.Row, 22)))
                MS.Text = Int(Val(M_FLX_AMT))
                MS.TextMatrix(MS.Row, 11) = Val(MS.TextMatrix(MS.Row, 3)) - (Val(MS.TextMatrix(MS.Row, 5)) + Val(MS.TextMatrix(MS.Row, 7)) + Val(MS.TextMatrix(MS.Row, 9)))
                MS.TextMatrix(MS.Row, 12) = QTY(STM(Val(MTS(Val(MS.TextMatrix(MS.Row, 4)), Val(sunt)) - (MTS(Val(MS.TextMatrix(MS.Row, 6)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 8)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 10)), Val(sunt)))), Val(sunt)))
                MS.Col = MS.Col + 1
                M_FLX_AMT.Visible = False
                MS.SetFocus
            End If
        'End If
    Else
        MS.Text = convert(M_FLX_AMT)
        M_FLX_AMT.Visible = False
        MS.SetFocus
    End If
End If
End Sub


Private Sub MS_EnterCell()
If MS.Col = 0 Then
    M_MDI.STS_BAR.Panels(1).Text = "Voucher Number"
ElseIf MS.Col = 1 Then
    M_MDI.STS_BAR.Panels(1).Text = "Short Name"
ElseIf MS.Col = 2 Then
    M_MDI.STS_BAR.Panels(1).Text = "Product Name"
ElseIf MS.Col = 3 Then
    M_MDI.STS_BAR.Panels(1).Text = "Total Shells"
ElseIf MS.Col = 4 Then
    M_MDI.STS_BAR.Panels(1).Text = "Total Bottles"
ElseIf MS.Col = 5 Then
    M_MDI.STS_BAR.Panels(1).Text = "Previous Received Shells"
ElseIf MS.Col = 6 Then
    M_MDI.STS_BAR.Panels(1).Text = "Previous Received Bottles"
ElseIf MS.Col = 7 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Current Receivable Shells"
ElseIf MS.Col = 8 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Current Receivable Bottles"
ElseIf MS.Col = 9 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Breakage Shells"
ElseIf MS.Col = 10 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Breakage Bottles"
ElseIf MS.Col = 11 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Diffrence Shells"
ElseIf MS.Col = 12 Then
    M_MDI.STS_BAR.Panels(1).Text = "Enter Diffrence Bottles"
ElseIf MS.Col = 13 Then
    M_MDI.STS_BAR.Panels(1).Text = "Deposit Shells"
ElseIf MS.Col = 14 Then
    M_MDI.STS_BAR.Panels(1).Text = "Deposit Bottles"
ElseIf MS.Col = 15 Then
    M_MDI.STS_BAR.Panels(1).Text = "Deposit Amount"
End If
End Sub


Private Sub MS_KeyPress(K As Integer)
If MS.Col <> 1 Then If Trim(MS.TextMatrix(MS.Row, 1)) = "" Then Exit Sub
'If cmbTrans.ListIndex = 0 Or cmbTrans.ListIndex = 1 Or cmbTrans.ListIndex = 9 Or cmbTrans.ListIndex = 10 Or cmbTrans.ListIndex = 11 Then
    If (MS.Col = 3 Or MS.Col = 4) And MS.Row > 0 Then
        If K > 47 And K < 58 Or K = 46 Then
            M_FLX_AMT = ""
            M_FLX_AMT.MaxLength = 7
            M_FLX_AMT.Visible = True
            M_FLX_AMT.Top = MS.Top + MS.CellTop
            M_FLX_AMT.Left = MS.CellLeft + MS.Left
            M_FLX_AMT.Width = MS.CellWidth
            M_FLX_AMT.Height = MS.CellHeight
            M_FLX_AMT.Visible = True
            M_FLX_AMT.SetFocus
            If K = vbKeySubtract Then
                K = 0
            Else
                M_FLX_AMT = Chr(K)
                M_FLX_AMT.SelStart = Len(M_FLX_AMT)
            End If
        ElseIf K = 13 Then
            If MS.Col = 4 Then
                If MS.Row + 1 < MS.Rows Then MS.Row = MS.Row + 1
                MS.Col = 3
            ElseIf MS.Col = 3 Then
                MS.Col = 4
            End If
        End If
    ElseIf MS.Col = 1 Then
        Text3 = ""
        If Trim(MS.Text) <> "" Then Exit Sub
        If K = 13 Then
            'DISPLIST
        ElseIf K <> 27 Then
            If K = vbKeySpace Then
            Else
                If Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
                Else
                    Text3.Visible = True
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
    ElseIf ((MS.Col >= 7 And MS.Col <= 12) Or MS.Col = 15) And MS.Row > 0 And Val(MS.Text) >= 0 Then
        If K > 47 And K < 58 Or K = 46 Then
            M_FLX_AMT = ""
            M_FLX_AMT.MaxLength = 7
            M_FLX_AMT.Visible = True
            M_FLX_AMT.Top = MS.Top + MS.CellTop
            M_FLX_AMT.Left = MS.CellLeft + MS.Left
            M_FLX_AMT.Width = MS.CellWidth
            M_FLX_AMT.Height = MS.CellHeight
            M_FLX_AMT.Visible = True
            M_FLX_AMT.SetFocus
            If K = vbKeySubtract Then
            Else
                M_FLX_AMT = Chr(K)
                M_FLX_AMT.SelStart = Len(M_FLX_AMT)
            End If
        ElseIf K = 13 Then
            If MS.Col = 8 Then
                If MS.Row + 1 < MS.Rows Then MS.Row = MS.Row + 1
                MS.Col = 7
            ElseIf MS.Col = 7 Then
                MS.Col = 8
            End If
        End If
    ElseIf MS.Col = 1 Then
        Text3 = ""
        If K = 13 Then
            'DISPLIST
        ElseIf K <> 27 Then
            If K = vbKeySpace Then
            Else
                If Trim(MS.TextMatrix(MS.Row, 0)) <> "" Then
                Else
                    Text3.Visible = True
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
    End If
'End If
End Sub


Private Sub MS_LeaveCell()
Text3.Visible = False
M_FLX_AMT.Visible = False
End Sub

Private Sub MS_RowColChange()
lbl_ret = Trim(MS.TextMatrix(MS.Row, 20))
End Sub


Private Sub MS_Scroll()
If Text3.Visible = True Then
    Text3 = ""
    Text3.Visible = False
    MS.SetFocus
ElseIf M_FLX_AMT.Visible = True Then
    M_FLX_AMT = ""
    M_FLX_AMT.Visible = False
    MS.SetFocus
ElseIf txtbrk.Visible = True Then
    txtbrk = ""
    txtbrk.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "s" Then
    saverecord
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If M_FRM_TRS.ActiveControl.name = "invid" Then V_INV_LST
ElseIf Button.KEY = "p" Then
    If CHECKTB("prt.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If
    If check("prt", "prt_nme", "M_FRM_TRS") = False Then
        M_FRM_PRV.Check3.Value = 0
    Else
        M_FRM_PRV.Check3.Value = 1
    End If
    M_FRM_PRV.Tag = "TRS"
    M_FRM_PRV.Show 1
    'Toolbar1.Enabled = False
    'SSTab1.Enabled = False
    End If
End Sub
Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_TRS.ActiveControl Is TextBox Then M_FRM_TRS.ActiveControl.Text = Trim(FUNKEY(K))
If K = 83 And s = 2 Then
    saverecord
ElseIf K = 81 And s = 2 Then
    cancel
ElseIf K = 88 And s = 2 Then
    Unload Me
ElseIf K = 69 And s = 2 Then
    If M_FRM_TRS.ActiveControl.name = "invid" Then V_INV_LST
ElseIf K = 112 Then
    M_HELP
End If
End Sub

Public Sub displaydetails(invid1 As String, txntyp As String)
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
rs.Open "select distinct SNV.INV_IDY,SNV.INV_IDY,SNV.PRD_PCK,SNV.PRD_UPC,SUM(I_SHL_NOS),SUM(I_PCK_NOS),SUM(R_SHL_NOS),SUM(R_PCK_NOS),SUM(BRK_SHL), SUM(BRK_PCK), SUM(D_SHL_NOS), SUM(D_PCK_NOS), SUM(DP_SHL_NOS), SUM(DP_PCK_NOS), SUM(DEP_AMT), SNV.TXN_TYP, SNV.INV_DAT,SNV.SHL_IDY, SNV.PRD_IDY from snv where inv_idy like '" & Trim(invid1) & "' and txn_typ like '" & Trim(txntyp) & "' Group by snv.prd_pck,SNV.PRD_UPC,snv.SHL_IDY", con, adOpenKeyset, adLockPessimistic
'RS.Open "select distinct SNV.INV_IDY,SNV.INV_IDY,SNV.PRD_PCK,SNV.PRD_UPC,SUM(I_SHL_NOS),SUM(I_PCK_NOS),SUM(R_SHL_NOS),SUM(R_PCK_NOS),SUM(BRK_SHL), SUM(BRK_PCK), SUM(D_SHL_NOS), SUM(D_PCK_NOS), SUM(DP_SHL_NOS), SUM(DP_PCK_NOS), SUM(DEP_AMT), SNV.TXN_TYP, SNV.INV_DAT,SNV.SHL_IDY from snv where inv_idy like '" & Trim(invid1) & "' and txn_typ like '" & Trim(txntyp) & "' Group by snv.prd_pck,SNV.PRD_UPC", CON, adOpenKeyset, adlockpessimistic 'Modified by Ramanesh - Because of Expected Query Error!
If rs.RecordCount > 0 Then
    invid = Trim(invid1)
    invid.Enabled = False
    If RS1.State = 1 Then RS1.Close
    RS1.Open "select * from inl,act where inl.cst_idy like act.act_idy and inl.inv_idy like '" & Trim(rs!INV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If RS1.RecordCount > 0 Then
        Label8 = Len(Trim(RS1!ORD_IDY))
        Label7 = RS1!CST_IDY
        Label4 = Trim(RS1!ACT_NME)
        Label5 = Trim(RS1!act_ad1)
        Label6 = Trim(RS1!act_ad4)
    End If
    cmbTrans.Enabled = False
    lv.Visible = False
    SSTab1.Tab = 0
End If
For A = 1 To rs.RecordCount
                    'MS.TextMatrix(A, 3) = Trim(RS!i_shl_nos)       'code modified by Ramesh Babu for PACK SIZE
                    'MS.TextMatrix(A, 4) = QTY(STM(Val(RS!i_pck_nos), Val(RS("SUB_UNT"))))
                    'MS.TextMatrix(A, 7) = Trim(RS!R_shl_nos & "")
                    'MS.TextMatrix(A, 8) = QTY(STM(Val(RS!R_pck_nos & ""), Val(RS("SUB_UNT"))))
                    'MS.TextMatrix(A, 9) = Trim(RS!brk_shl)
                    'MS.TextMatrix(A, 10) = QTY(STM(Val(RS!brk_pck), Val(RS("SUB_UNT"))))
                    'MS.TextMatrix(A, 11) = Trim(RS!D_shl_nos & "")
                    'MS.TextMatrix(A, 12) = QTY(STM(Val(RS!D_pck_nos & ""), Val(RS("SUB_UNT"))))
                    'MS.TextMatrix(A, 13) = Trim(RS!dp_shl_nos & "")
                    'MS.TextMatrix(A, 14) = QTY(STM(Val(RS!dp_pck_nos), Val(RS("SUB_UNT"))))
                    'MS.TextMatrix(A, 15) = convert(Val(RS!dep_amt & ""))
    MS.AddItem ""
    MS.TextMatrix(A, 0) = Trim(rs!INV_IDY)
    
    MS.TextMatrix(A, 1) = GetValue("SHL_NME", "SHL WHERE SHL_IDY LIKE '" & Trim(rs!shl_idy) & "'")
    MS.TextMatrix(A, 2) = Trim(rs!PRD_PCK & "") & " - " & Val(rs!prd_upc) & ""
    MS.TextMatrix(A, 3) = Val(rs(4))
    MS.TextMatrix(A, 4) = QTY(STM(Val(rs(5)), Val(rs("PRD_UPC"))))
    MS.TextMatrix(A, 7) = Val(rs(6))
    MS.TextMatrix(A, 8) = QTY(STM(Val(rs(7)), Val(rs("PRD_UPC"))))
    MS.TextMatrix(A, 9) = Val(rs(8))
    MS.TextMatrix(A, 10) = QTY(STM(Val(rs(9)), Val(rs("PRD_UPC"))))
    MS.TextMatrix(A, 11) = Val(rs(10))
    MS.TextMatrix(A, 12) = QTY(STM(Val(rs(11)), Val(rs("PRD_UPC"))))
    MS.TextMatrix(A, 13) = Val(rs(12))
    MS.TextMatrix(A, 14) = QTY(STM(Val(rs(13)), Val(rs("PRD_UPC"))))
    MS.TextMatrix(A, 15) = Val(rs(14))
    MS.TextMatrix(A, 16) = Trim(rs!prd_idy)
    MS.TextMatrix(A, 17) = Trim(rs!shl_idy)
    MS.TextMatrix(A, 18) = Format(rs!INV_DAT, "DD/MMM/YYYY")
    MS.TextMatrix(A, 19) = Trim(rs!TXN_TYP)
    MS.TextMatrix(A, 22) = Trim(rs!prd_upc)
    MS.TextMatrix(A, 23) = Trim(rs!PRD_PCK & "")
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
rs.Open "SELECT SUB_UNT,SHL_IDY,PRD_PCK,SHL_NME,PRD_PCK FROM PRD,SHL WHER PRD.PRD_IDY LIKE SHL.PRD_IDY GROUP BY SUB_UNT,SHL_IDY,PRD_PCK", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If Trim(rs!PRD_PCK) = Trim(MS.TextMatrix(b, 23)) And Trim(rs!shl_idy) = Trim(MS.TextMatrix(b, 17)) And Val(rs!SUB_UNT) = Val(MS.TextMatrix(b, 22)) Then
            GoTo o:
        End If
    Next
    MS.AddItem ""
    MS.TextMatrix(MS.Rows - 1, 0) = Trim(invid)
    MS.TextMatrix(MS.Rows - 1, 1) = GetValue("SHL_NME", "SHL WHERE SHL_IDY LIKE '" & Trim(rs!shl_idy) & "'")
    MS.TextMatrix(MS.Rows - 1, 2) = Trim(rs!PRD_PCK & "") & " - " & Val(rs!SUB_UNT) & ""
'    MS.TextMatrix(A, 3) =
 '   MS.TextMatrix(A, 4) = QTY(STM(Val(RS(5)), Val(RS("PRD_UPC"))))
 '   MS.TextMatrix(A, 7) = Val(RS(6))
 '   MS.TextMatrix(A, 8) = QTY(STM(Val(RS(7)), Val(RS("PRD_UPC"))))
 '   MS.TextMatrix(A, 9) = Val(RS(8))
 '   MS.TextMatrix(A, 10) = QTY(STM(Val(RS(9)), Val(RS("PRD_UPC"))))
 '   MS.TextMatrix(A, 11) = Val(RS(10))
 '   MS.TextMatrix(A, 12) = QTY(STM(Val(RS(11)), Val(RS("PRD_UPC"))))
 '   MS.TextMatrix(A, 13) = Val(RS(12))
 '   MS.TextMatrix(A, 14) = QTY(STM(Val(RS(13)), Val(RS("PRD_UPC"))))
  '  MS.TextMatrix(A, 15) = Val(RS(14))
   ' MS.TextMatrix(A, 16) = Trim(RS!PRD_IDY)
    MS.TextMatrix(MS.Rows - 1, 17) = Trim(rs!shl_idy)
  '  MS.TextMatrix(A, 18) = Format(RS!INV_DAT, "DD/MMM/YYYY")
    MS.TextMatrix(MS.Rows - 1, 19) = Trim(txntyp)
    MS.TextMatrix(MS.Rows - 1, 22) = Trim(rs!SUB_UNT)
    MS.TextMatrix(MS.Rows - 1, 23) = Trim(rs!PRD_PCK & "")
o:
    rs.MoveNext
Next
If Trim(MS.TextMatrix(MS.Rows - 1, 0)) = "" Then
    MS.Rows = MS.Rows - 1
End If
'MS.SetFocus
End Sub
Public Sub cancel()
clrctr Me
MS.Rows = 1
invid.Enabled = True
cmbTrans.Enabled = True
Label4 = ""
Label5 = ""
Label6 = ""
cmbTrans.SetFocus
End Sub
Public Sub saverecord()
Dim SMNIDY, ROUIDY As String
If MsgBox("Confirm Transaction Save ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbNo Then
cancel
Exit Sub
End If
If Trim(invid) <> "" And invid.Enabled = True Then
    MsgBox "Select Invoice ID"
    If invid.Enabled = True Then
        invid.SetFocus
        SLT
    Else
        invid.SetFocus
    End If
Else
    TxnMoveBackup "SNV", "SBK", Trim(invid), Trim(ttype)
    con.Execute "DELETE FROM SNV WHERE INV_IDY LIKE '" & Trim(invid) & "' AND TXN_TYP LIKE '" & ttype & "'"
    Dim RSSNV As New ADODB.Recordset
    RSSNV.Open "SELECT * FROM SNV WHERE INV_IDY LIKE ''", con, adOpenKeyset, adLockPessimistic
    STKUPD = GetValue("STK_UPD", "INV WHERE INV_IDY LIKE '" & Trim(invid) & "'")
    SMNIDY = GetValue("SMN_IDY", "INL WHERE INV_IDY LIKE '" & Trim(invid) & "'")
    ROUIDY = GetValue("ROU_IDY", "INL WHERE INV_IDY LIKE '" & Trim(invid) & "'")
    For A = 1 To MS.Rows - 1
        If MS.TextMatrix(A, 1) <> "" Then
            RSSNV.AddNew
            RSSNV!INV_IDY = Trim(invid)
            RSSNV!INV_DAT = GetValue("INV_DAT", "INL WHERE INV_IDY LIKE '" & Trim(invid) & "'")
            RSSNV!prd_idy = MS.TextMatrix(A, 16)
            RSSNV!TXN_TYP = Trim(ttype)
            RSSNV!shl_idy = Trim(MS.TextMatrix(A, 17))
            M_SUB_UNT = Val(MS.TextMatrix(A, 22))
            RSSNV!pck_idy = ""
            RSSNV!I_shl_nos = Val(MS.TextMatrix(A, 3))
            RSSNV!I_pck_nos = MTS(Val(MS.TextMatrix(A, 4)), Val(M_SUB_UNT))
            RSSNV!R_shl_nos = Val(MS.TextMatrix(A, 7))
            RSSNV!R_pck_nos = MTS(Val(MS.TextMatrix(A, 8)), Val(M_SUB_UNT))
            RSSNV!d_shl_nos = Val(Trim(MS.TextMatrix(A, 11)) & "")
            RSSNV!d_pck_nos = MTS(Val(MS.TextMatrix(A, 12)), Val(M_SUB_UNT))
            RSSNV!dp_shl_nos = Val(Trim(MS.TextMatrix(A, 13)) & "")
            RSSNV!dp_pck_nos = MTS(Val(MS.TextMatrix(A, 14)), Val(M_SUB_UNT))
            RSSNV!dep_amt = Val(Trim(MS.TextMatrix(A, 15)) & "")
            RSSNV!BRK_SHL = Val(MS.TextMatrix(A, 9))
            RSSNV!BRK_PCK = MTS(Val(MS.TextMatrix(A, 10)), Val(M_SUB_UNT))
            RSSNV!txn_dat = Trim(RSSNV!INV_DAT)
            If ttype = "SAL" Then
                RSSNV!TXN_IDY = GetValue("ord_idy", "inl where inv_idy like '" & Trim(invid) & "'")
            Else
                RSSNV!TXN_IDY = ""
            End If
            RSSNV!SMN_IDY = Trim(SMNIDY)    'CHANGED BY PKB
            RSSNV!rou_idy = Trim(ROUIDY)     'CHANGED BY PKB
            RSSNV!CST_IDY = Trim(Label7)
            RSSNV!PRD_PCK = Trim(MS.TextMatrix(A, 23))
            RSSNV!prd_upc = Val(MS.TextMatrix(A, 22))
            RSSNV!stk_upd = STKUPD
            RSSNV.Update
        End If
    Next
    If ttype = "PUR" Or ttype = "SRN" Or ttype = "STI" Then If check("PRT", "PRT_NME", "TRS_SHL") = True Then LOADSHELLSSORTING ttype, Trim(invid)
    cancel
End If
End Sub

Public Sub calc()
For A = 1 To MS.Rows - 1
MS.Row = A
    If MS.TextMatrix(A, 1) <> "" Then
        For b = 7 To 8
            MS.Col = b
            If MS.Col = 9 Or MS.Col = 10 Then
                MS.TextMatrix(MS.Row, 11) = Val(MS.TextMatrix(MS.Row, 3)) - (Val(MS.TextMatrix(MS.Row, 5)) + Val(MS.TextMatrix(MS.Row, 7)) + Val(MS.TextMatrix(MS.Row, 9)))
                MS.TextMatrix(MS.Row, 12) = QTY(STM(Val(MTS(Val(MS.TextMatrix(MS.Row, 4)), Val(sunt)) - (MTS(Val(MS.TextMatrix(MS.Row, 6)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 8)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 10)), Val(sunt)))), Val(sunt))) ' Val(MS.TextMatrix(MS.Row, 4)) - (Val(MS.TextMatrix(MS.Row, 6)) + Val(MS.TextMatrix(MS.Row, 8)) + Val(MS.TextMatrix(MS.Row, 10)))
            ElseIf MS.Col = 7 Or MS.Col = 8 Then
                If MS.Col = 7 Then MS.TextMatrix(MS.Row, 8) = MTS(Val(MS.Text), Val(MS.TextMatrix(MS.Row, 22)))
                MS.TextMatrix(MS.Row, 11) = Val(MS.TextMatrix(MS.Row, 3)) - (Val(MS.TextMatrix(MS.Row, 5)) + Val(MS.TextMatrix(MS.Row, 7)) + Val(MS.TextMatrix(MS.Row, 9)))
                MS.TextMatrix(MS.Row, 12) = QTY(STM(Val(MTS(Val(MS.TextMatrix(MS.Row, 4)), Val(sunt)) - (MTS(Val(MS.TextMatrix(MS.Row, 6)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 8)), Val(sunt)) + MTS(Val(MS.TextMatrix(MS.Row, 10)), Val(sunt)))), Val(sunt))) 'Val(MS.TextMatrix(MS.Row, 4)) - (Val(MS.TextMatrix(MS.Row, 6)) + Val(MS.TextMatrix(MS.Row, 8)) + Val(MS.TextMatrix(MS.Row, 10)))
                If cmbTrans.ListIndex = 3 Or cmbTrans.ListIndex = 4 Or cmbTrans.ListIndex = 9 Or cmbTrans.ListIndex = 10 Or cmbTrans.ListIndex = 11 Then
                    'If Trim(MS.TextMatrix(MS.Rows - 1, 1)) <> "" Then MS.AddItem ""
                    If MS.Col = 8 Then
                        If MS.Row + 1 < MS.Rows Then MS.Row = MS.Row + 1
                        MS.Col = 7
                    ElseIf MS.Col = 7 Then
                        MS.Col = 8
                    End If
                End If
            ElseIf MS.Col = 11 Or MS.Col = 12 Then
                MS.TextMatrix(MS.Row, 13) = Val(MS.TextMatrix(MS.Row, 3)) - (Val(MS.TextMatrix(MS.Row, 5)) + Val(MS.TextMatrix(MS.Row, 7)) + Val(MS.TextMatrix(MS.Row, 9)) + Val(MS.TextMatrix(MS.Row, 11)))
                MS.TextMatrix(MS.Row, 14) = QTY(Val(MS.TextMatrix(MS.Row, 4)) - (Val(MS.TextMatrix(MS.Row, 6)) + Val(MS.TextMatrix(MS.Row, 8)) + Val(MS.TextMatrix(MS.Row, 10)) + Val(MS.TextMatrix(MS.Row, 12))))
            ElseIf MS.Col = 3 Or MS.Col = 4 Then
                If cmbTrans.ListIndex = 3 Or cmbTrans.ListIndex = 4 Or cmbTrans.ListIndex = 9 Or cmbTrans.ListIndex = 10 Or cmbTrans.ListIndex = 11 Then
                    MS.Text = Int(Val(M_FLX_AMT))
                    'If Trim(MS.TextMatrix(MS.Rows - 1, 1)) <> "" Then MS.AddItem ""
                    If MS.Col = 4 Then
                        If MS.Row + 1 < MS.Rows Then MS.Row = MS.Row + 1
                        MS.Col = 3
                    ElseIf MS.Col = 3 Then
                        MS.TextMatrix(MS.Row, 4) = MTS(Val(MS.Text), Val(MS.TextMatrix(MS.Row, 22)))
                        MS.Col = MS.Col + 1
                    End If
                End If
            Else
                MS.Text = convert(M_FLX_AMT)
            End If
        Next
    End If
Next
End Sub

Public Function CSUBUNT(STK As Double, SUBUNT As Integer) As Boolean
S_UNT = (Val(STK) - Int(STK)) * 1000
If S_UNT < SUBUNT Then
    CSUBUNT = True
Else
    MsgBox "Product Subunits are " & SUBUNT, vbExclamation
    M_FLX_AMT.SetFocus
    SendKeys "{HOME}+{END}"
    CSUBUNT = False
End If
End Function

Public Sub loaddetails()
STR1 = Split(Gltxninvid, ",")
Gltxninvid = ""
If STR1(0) = "SAL" Then
    cmbTrans.ListIndex = 1
ElseIf STR1(0) = "PUR" Then
    cmbTrans.ListIndex = 0
ElseIf STR1(0) = "SRN" Then
    cmbTrans.ListIndex = 2
ElseIf STR1(0) = "PRT" Then
    cmbTrans.ListIndex = 3
ElseIf STR1(0) = "STI" Then
    cmbTrans.ListIndex = 4
ElseIf STR1(0) = "STO" Then
    cmbTrans.ListIndex = 5
End If
    displaydetails Trim(STR1(1)), Trim(STR1(0))
End Sub
