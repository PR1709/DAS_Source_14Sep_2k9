VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form M_FRM_RSI_PRD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Van Stock Issue SKU Quick Select"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "M_FRM_RSI_PRD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Select All"
      Height          =   450
      Left            =   3720
      TabIndex        =   5
      Top             =   4940
      Width           =   1050
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Deselect All"
      Height          =   450
      Left            =   5040
      TabIndex        =   4
      Top             =   4940
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   6240
      TabIndex        =   1
      Top             =   4700
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   450
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1050
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   450
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1050
      End
   End
   Begin MSComctlLib.ListView list1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   8830
      _ExtentX        =   15584
      _ExtentY        =   8281
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
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product (SKU) Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Short Name"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Sale Rate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "MRP"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Quantity"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "UPC"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Cnv.Fac"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "UOM"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Segment"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Brand Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Supplier Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Suplr.Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_RSI_PRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
For A = 1 To list1.ListItems.COUNT
    If list1.ListItems.Item(A).Checked = True Then
    PrdCollection.Add Trim(list1.ListItems.Item(A).ListSubItems(13))
    End If
Next
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
For A = 1 To list1.ListItems.COUNT
    list1.ListItems.Item(A).Checked = True
Next
End Sub

Private Sub Command4_Click()
For A = 1 To list1.ListItems.COUNT
    list1.ListItems.Item(A).Checked = False
Next
End Sub

Private Sub Form_Load()
PRDLIST
End Sub

Private Sub PRDLIST()
If CHECKTB("act.dbf,pgr.dbf,prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select prd.prd_idy,prd.prd_nme,prd.prd_prn,prd.brd_nme,prd.alt_idy,brd.prd_sdr,brd.prd_mrp,brd.prd_qty,act.act_nme,pgr.grp_cat,prd.prd_uom,prd.sub_unt,prd.con_fac,bat_idy from brd,prd,act,pgr where brd.prd_idy like prd.prd_idy and act.act_idy like prd.cmp_idy  and pgr.grp_idy like prd.grp_idy and brd.prd_qty>0", CON, adOpenKeyset, adLockPessimistic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("PRD_NME")))
        ls.ListSubItems.Add , , Trim(rs("PRD_PRN"))
        ls.ListSubItems.Add , , AMT(Trim(rs("PRD_SDR")))
        ls.ListSubItems.Add , , AMT(Trim(rs("PRD_MRP")))
        ls.ListSubItems.Add , , Trim(QTY(rs("prd_qty") & ""))
        ls.ListSubItems.Add , , Trim(rs("SUB_UNT") & "")
        ls.ListSubItems.Add , , QTY(rs("CON_FAC") & "")
        ls.ListSubItems.Add , , Trim(rs("PRD_UOM") & "")
        ls.ListSubItems.Add , , Trim(rs("GRP_CAT") & "")
        ls.ListSubItems.Add , , Trim(rs("BRD_NME") & "")
        ls.ListSubItems.Add , , Trim(rs("ACT_NME") & "")
        ls.ListSubItems.Add , , Trim(rs("ALT_IDY") & "")
        ls.ListSubItems.Add , , Trim(rs("PRD_IDY") & "")
        ls.ListSubItems.Add , , Trim(rs("BAT_IDY") & "")
        rs.MoveNext
    Next
    list1.Visible = True
End If
End Sub

