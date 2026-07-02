VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_XRF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reference Documents"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   Icon            =   "M_FRM_XRF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4850
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtVanRef 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtMERNum 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      MaxLength       =   24
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2535
   End
   Begin MSComctlLib.ListView lvVanRef 
      Height          =   1575
      Left            =   1680
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Van Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Salesman"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Route"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView lvMERNum 
      Height          =   1575
      Left            =   1680
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2760
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MER Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Scheme Name"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Van Challan ID"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   210
      TabIndex        =   6
      Top             =   210
      Width           =   1065
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MER Reference"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1155
   End
End
Attribute VB_Name = "M_FRM_XRF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GF As Integer

Private Sub Command1_Click()
VanRef = txtVanRef.Text
MerNum = txtMERNum.Text
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
    DisplayAdditional lvVanRef, lvMERNum, txtVanRef, txtMERNum
    txtMERNum.Text = MerNum
    txtVanRef.Text = VanRef
    DisplayValues
End Sub
Private Sub Form_Unload(Cancel As Integer)
VanRef = txtVanRef.Text
MerNum = txtMERNum.Text

End Sub

Private Sub txtMERNum_Change()
    Dim status As Boolean
    status = False
    For i = 1 To lvMERNum.ListItems.count
        If Mid(Trim(UCase(lvMERNum.ListItems(i).Text)), 1, Len(txtMERNum.Text)) = UCase(txtMERNum.Text) Then
            lvMERNum.ListItems(i).Selected = True
            status = True
            Exit For
        End If
    Next
End Sub

Private Sub txtMERNum_GotFocus()
GF = 1
txtMERNum.SelStart = 0
txtMERNum.SelLength = Len(txtMERNum.Text)

End Sub

Private Sub txtMERNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        lvMERNum.SetFocus
    End If
End Sub


Private Sub lvVanRef_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Shift = vbCtrlMask Then
    lvVanRef.SelectedItem.Selected = False
    txtVanRef.Text = ""
ElseIf KeyCode = 13 Then
    txtVanRef.Text = Trim(lvVanRef.SelectedItem.Text)
    txtMERNum.SetFocus
End If
End Sub

Private Sub lvVanRef_ItemClick(ByVal Item As MSComctlLib.ListItem)
If GF = 0 Then txtVanRef.Text = Item.Text
End Sub



Public Sub DisplayAdditional(lvVanRef As ListView, lvMERNum As ListView, txtVanRef As TextBox, txtMERNum As TextBox)
Dim rs As New ADODB.Recordset
Dim Item As ListItem

lvVanRef.ListItems.CLEAR
txtVanRef.Text = ""
If rs.State = 1 Then rs.Close
rs.Open "SELECT DISTINCT INV_IDY , SMN_NME, ROU_IDY, VAN.SMN_IDY FROM VAN, SMN WHERE VAN.SMN_IDY =  SMN.SMN_IDY AND VAN_CLS='S'", con, adOpenKeyset, adLockReadOnly
While Not rs.EOF
    Set Item = lvVanRef.ListItems.Add(, , rs!INV_IDY)
    Item.ListSubItems.Add , , rs!SMN_NME
    Item.ListSubItems.Add , , rs!rou_idy
    rs.MoveNext
Wend
If lvVanRef.ListItems.count > 0 Then
    lvVanRef.ListItems(1).Selected = True
End If
lvMERNum.ListItems.CLEAR
txtMERNum.Text = ""
If rs.State = 1 Then rs.Close
rs.Open "SELECT DISTINCT MER_NUM, SCH_NME FROM SCH WHERE MER_NUM IS NOT NULL OR TRIM(MER_NUM) <> ''", con, adOpenKeyset, adLockReadOnly
While Not rs.EOF
    If Trim(rs!MER_NUM) <> "" Then
        Set Item = lvMERNum.ListItems.Add(, , rs!MER_NUM)
        Item.ListSubItems.Add , , rs!Sch_Nme
    End If
    rs.MoveNext
Wend
If lvMERNum.ListItems.count > 0 Then
    lvMERNum.ListItems(1).Selected = True
End If
If rs.State = 1 Then rs.Close

End Sub

Private Sub DisplayValues()

    If Trim(txtVanRef.Text) <> "" Then
    If lvVanRef.ListItems.count > 0 Then
        For i = 1 To lvVanRef.ListItems.count
            If Trim(lvVanRef.ListItems(i).Text) = Trim(txtVanRef.Text) Then
                lvVanRef.SelectedItem = lvVanRef.ListItems(i)
                Exit For
            End If
        Next
    End If
    End If
    If Trim(txtMERNum.Text) <> "" Then
    If lvMERNum.ListItems.count > 0 Then
        For i = 1 To lvMERNum.ListItems.count
            If Trim(lvMERNum.ListItems(i).Text) = Trim(txtMERNum.Text) Then
                lvMERNum.SelectedItem = lvMERNum.ListItems(i)
                Exit For
            End If
        Next
    End If
    End If
End Sub

Private Sub txtMERNum_LostFocus()
GF = 0
End Sub

Private Sub txtVanRef_Change()
    Dim status As Boolean
    status = False
    For i = 1 To lvVanRef.ListItems.count
        If Mid(Trim(UCase(lvVanRef.ListItems(i).Text)), 1, Len(txtVanRef.Text)) = UCase(txtVanRef.Text) Then
            lvVanRef.ListItems(i).Selected = True
            status = True
            Exit For
        End If
    Next
End Sub

Private Sub txtVanRef_GotFocus()
GF = 2
txtVanRef.SelStart = 0
txtVanRef.SelLength = Len(txtVanRef.Text)

End Sub

Private Sub txtVanRef_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
'        lvVanRef.SetFocus
    ElseIf KeyCode = 13 Then
        lvVanRef.SetFocus
    End If
End Sub

Private Sub txtVanRef_LostFocus()
GF = 0
End Sub


Private Sub lvMERNum_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Shift = vbCtrlMask Then
    lvMERNum.SelectedItem.Selected = False
ElseIf KeyCode = 13 Then
    txtMERNum.Text = Trim(lvMERNum.SelectedItem.Text)
    Command1.SetFocus
End If
End Sub

Private Sub lvMERNum_ItemClick(ByVal Item As MSComctlLib.ListItem)
If GF = 0 Then txtMERNum.Text = Item.Text
End Sub


