VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_REP_INT 
   Caption         =   "Retailer Sales Report"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   Icon            =   "M_REP_INT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5150
      Left            =   0
      TabIndex        =   1
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
      TabCaption(0)   =   "&Define Products"
      TabPicture(0)   =   "M_REP_INT.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LV2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FDATE"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TDATE"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbCat"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdGenerate"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "&View"
      TabPicture(1)   =   "M_REP_INT.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ms"
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid ms 
         Height          =   4695
         Left            =   -74940
         TabIndex        =   10
         Top             =   360
         Width           =   8650
         _ExtentX        =   15266
         _ExtentY        =   8281
         _Version        =   393216
         GridLines       =   2
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Generate"
         Height          =   375
         Left            =   7350
         TabIndex        =   9
         Top             =   1250
         Width           =   1335
      End
      Begin VB.ComboBox cmbCat 
         Height          =   315
         ItemData        =   "M_REP_INT.frx":047A
         Left            =   1320
         List            =   "M_REP_INT.frx":0490
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker TDATE 
         Height          =   330
         Left            =   3000
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24838147
         CurrentDate     =   37258
      End
      Begin MSComCtl2.DTPicker FDATE 
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24838147
         CurrentDate     =   37258
      End
      Begin MSComctlLib.ListView LV2 
         Height          =   3400
         Left            =   50
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1680
         Width           =   8650
         _ExtentX        =   15266
         _ExtentY        =   6006
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select"
         Height          =   195
         Left            =   720
         TabIndex        =   6
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   728
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   2640
         TabIndex        =   4
         Top             =   728
         Width           =   195
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
            Picture         =   "M_REP_INT.frx":04CE
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_REP_INT.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_REP_INT.frx":0DDA
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_REP_INT.frx":112E
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_REP_INT.frx":1766
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_REP_INT.frx":1ABA
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_REP_INT.frx":1F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_REP_INT.frx":220E
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_REP_INT.frx":261A
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_REP_INT.frx":2726
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_REP_INT.frx":2A7A
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_REP_INT.frx":2E8E
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
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "c"
            Object.ToolTipText     =   "Cancel (Ctrl+Q)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_REP_INT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSPRD As New ADODB.Recordset
Dim rsact As New ADODB.Recordset
Dim I As Integer, ls As ListItem, CNT As Integer
Dim sku(5) As String
Private Sub cmbCat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If cmbCat.ListIndex = 0 Then
        lvp
        PRDLIST
        LV2.SetFocus
    ElseIf cmbCat.ListIndex = 1 Then
        With LV2
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Pack Size", 2500
        End With
        packlist
        LV2.SetFocus
    ElseIf cmbCat.ListIndex = 2 Then
        With LV2
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Pack Type", 2500
        End With
        ptypelist
        LV2.SetFocus
    ElseIf cmbCat.ListIndex = 3 Then
        With LV2
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Brand Name", 2500
        End With
        brdlist
        LV2.SetFocus
    ElseIf cmbCat.ListIndex = 4 Then
        With LV2
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Segment", 2500
            .ColumnHeaders.Add , , "Group ID", 2500
        End With
        seglist
        LV2.SetFocus
    ElseIf cmbCat.ListIndex = 5 Then
        With LV2
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Category", 2500
            .ColumnHeaders.Add , , "Group ID", 2500
        End With
        catlist
        LV2.SetFocus
    End If
End If
End Sub

Private Sub cmdGenerate_Click()
SSTab1.Tab = 1
ms.Visible = True
If rsact.State = 1 Then rsact.Close
rsact.Open "select act_nme from act", CON, adOpenKeyset, adLockOptimistic
ms.Rows = rsact.RecordCount + 1
ms.Cols = CNT + 1
ms.TextMatrix(0, 1) = sku(0)
ms.TextMatrix(0, 2) = sku(1)
ms.TextMatrix(0, 3) = sku(2)
ms.TextMatrix(0, 4) = sku(3)
ms.TextMatrix(0, 5) = sku(4)
End Sub

Private Sub FDATE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    TDATE.SetFocus
End If
End Sub

Private Sub Form_Activate()
If FDATE.Enabled = True Then FDATE.SetFocus
End Sub

Private Sub Form_KeyDown(K As Integer, Shift As Integer)
If K = 120 Then Shell "c:\windows\calc.exe"
If K > 111 And K < 120 Then If TypeOf M_FRM_ACT.ActiveControl Is TextBox Then M_FRM_ACT.ActiveControl.Text = Trim(FUNKEY(K))
If K = 81 And S = 2 Then
    cancel
ElseIf K = 88 And S = 2 Then
    Unload Me
ElseIf K = 112 Then
    M_HELP
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
SETYEAR Me
cmbCat.ListIndex = 0
lvp
CNT = 0
End Sub
Private Sub lvp()
With LV2
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Product Name", 3500
    .ColumnHeaders.Add , , "Pack Size", 1700
    .ColumnHeaders.Add , , "Segment", 1800
    .ColumnHeaders.Add , , "Product ID", 1700
End With
End Sub

Private Sub LV2_Click()
If LV2.ListItems(1).Checked = False Then
    CNT = CNT - 1
Else
    If CNT < 5 Then
        CNT = CNT + 1
    Else
        MsgBox "Maximum No is 5", vbExclamation
        LV2.ListItems(LV2.SelectedItem.Index).Checked = False
        LV2.SetFocus
    End If
End If
End Sub

Private Sub LV2_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    If LV2.ListItems(LV2.SelectedItem.Index).Checked = False Then
        sku(CNT - 1) = ""
    Else
        If CNT < 5 Then
            CNT = CNT + 1
        Else
            MsgBox "Maximum No is 5", vbExclamation
            LV2.ListItems(LV2.SelectedItem.Index).Checked = False
            LV2.SetFocus
        End If
    End If
End If
End Sub

Private Sub TDATE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmbCat.SetFocus
End If
End Sub
Private Sub PRDLIST()
If RSPRD.State = 1 Then RSPRD.Close
RSPRD.Open "SELECT PRD.PRD_IDY,PRD.PRD_NME,PRD.PRD_PCK,PRD.GRP_IDY,PGR.GRP_CLS AS PRD_SEG FROM PRD,PGR WHERE PRD.GRP_IDY LIKE PGR.GRP_IDY", CON, adOpenKeyset, adLockOptimistic
For I = 0 To RSPRD.RecordCount - 1
    Set ls = LV2.ListItems.Add(, , RSPRD("PRD_NME"))
    ls.ListSubItems.Add , , RSPRD("PRD_PCK")
    ls.ListSubItems.Add , , RSPRD("PRD_SEG")
    ls.ListSubItems.Add , , RSPRD("PRD_IDY")
    RSPRD.MoveNext
Next
End Sub
Private Sub cancel()
FDATE.Value = Date
TDATE.Value = Date
LV2.ListItems.Clear
cmbCat.ListIndex = 0
SSTab1.Tab = 0
FDATE.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "p" Then

ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "e" Then
    Unload Me
End If
End Sub
Private Sub packlist()
If RSPRD.State = 1 Then RSPRD.Close
RSPRD.Open "select prd_pck from prd group by prd_pck", CON, adOpenKeyset, adLockOptimistic
LV2.ListItems.Clear
For I = 0 To RSPRD.RecordCount - 1
    Set ls = LV2.ListItems.Add(, , RSPRD("prd_pck"))
    RSPRD.MoveNext
Next
End Sub
Private Sub ptypelist()
If RSPRD.State = 1 Then RSPRD.Close
RSPRD.Open "select prd_pkt from prd group by prd_pkt", CON, adOpenKeyset, adLockOptimistic
LV2.ListItems.Clear
For I = 0 To RSPRD.RecordCount - 1
    Set ls = LV2.ListItems.Add(, , RSPRD("prd_pkt"))
    RSPRD.MoveNext
Next
End Sub

Private Sub brdlist()
If RSPRD.State = 1 Then RSPRD.Close
LV2.ListItems.Clear
RSPRD.Open "select brd_nme from prd group by brd_nme", CON, adOpenKeyset, adLockOptimistic
For I = 0 To RSPRD.RecordCount - 1
    Set ls = LV2.ListItems.Add(, , RSPRD("brd_nme"))
    RSPRD.MoveNext
Next
End Sub
Private Sub seglist()
If RSPRD.State = 1 Then RSPRD.Close
RSPRD.Open "select grp_idy,grp_cls from pgr group by grp_cls", CON, adOpenKeyset, adLockOptimistic
LV2.ListItems.Clear
For I = 0 To RSPRD.RecordCount - 1
    Set ls = LV2.ListItems.Add(, , RSPRD("grp_cls"))
    ls.ListSubItems.Add , , RSPRD("grp_idy")
    RSPRD.MoveNext
Next
End Sub
Private Sub catlist()
If RSPRD.State = 1 Then RSPRD.Close
RSPRD.Open "select grp_idy,grp_cat from pgr group  by grp_cat", CON, adOpenKeyset, adLockOptimistic
LV2.ListItems.Clear
For I = 0 To RSPRD.RecordCount - 1
    Set ls = LV2.ListItems.Add(, , RSPRD("grp_cat"))
    ls.ListSubItems.Add , , RSPRD("grp_idy")
    RSPRD.MoveNext
Next
End Sub
