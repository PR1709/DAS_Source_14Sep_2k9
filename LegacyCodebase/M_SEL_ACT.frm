VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_SEL_ACT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Sales Account"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "M_SEL_ACT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView list2 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Sales Account View"
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
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "M_SEL_ACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub V_ACT_LST()
End Sub

Private Sub Form_Load()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
typ = ""
If typ = "" Then
typ = Me.Tag
End If
REC_SET.Open "select * from act where grp_clf like 'Sales' and ACT_TYP LIKE '' ", CON, adOpenKeyset, adLockOptimistic
List2.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = List2.ListItems.Add(, , Trim(REC_SET("act_idy")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_nme"))
        REC_SET.MoveNext
    Next
Else
End If
End Sub

Private Sub list2_DblClick()
list2_keypress 13
End Sub

Private Sub list2_keypress(k As Integer)
If k = 13 Then
    'M_FRM_RSI.Tag = List2.SelectedItem
    sal_act = List2.SelectedItem
    Unload M_SEL_ACT
End If
End Sub
