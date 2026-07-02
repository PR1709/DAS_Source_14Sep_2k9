VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_RQY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Filters"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "M_FRM_RQY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4048
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Report Filters"
      TabPicture(0)   =   "M_FRM_RQY.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Check1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_RQY.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "LIST1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   3480
         TabIndex        =   7
         Top             =   480
         Width           =   1215
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView LIST1 
         Height          =   1215
         Left            =   -74640
         TabIndex        =   4
         Top             =   840
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   2143
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Sales Man"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&All"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1492
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   3360
         TabIndex        =   6
         Top             =   1800
         Width           =   45
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Supplier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   -74640
         TabIndex        =   5
         Top             =   480
         Width           =   4305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Issue Challan No"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   1560
         Width           =   1200
      End
   End
End
Attribute VB_Name = "M_FRM_RQY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Command1_Click
ElseIf Check1.Value = 0 Then
    Check1.Enabled = False
    Text1.SetFocus
End If
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Check1.Value = 1 Then
    SAL_QRY = "(TRIM(SUBSTR(INL->INV_IDY,1,3))<>'SAL').AND.INL->TXN_TYP='SAL'"
Else
    SAL_QRY = "INL->ORD_IDY='" & Trim(Text1) & "'.AND.(TRIM(SUBSTR(INL->INV_IDY,1,3))<>'SAL').AND.INL->TXN_TYP='SAL'"
    'SAL_QRY = "INL->ORD_IDY='" & Trim(Text1) & "'"
End If
Openreport CURDIR & "\report.rp1", "Sales Report - Van Challan", 2, SAL_QRY, 6
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DTPicker1_KeyDown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then Text1.SetFocus
End Sub

Private Sub Form_Activate()
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
Check1.SetFocus
SETYEAR M_FRM_RQY
End Sub

Private Sub Form_Keypress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        Text1.SetFocus
    Else
        Unload Me
    End If
End If
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

Private Sub List1_keypress(k As Integer)
If k = 13 Then
    Text1 = Trim(list1.SelectedItem)
    Check1.Value = 0
    Check1.Enabled = False
    list1.Visible = False
    SSTab1.Tab = 0
    Command1.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(k As Integer)
If k = 13 Then
    If Len(Trim(Text1)) = 0 Then
        Supplierlist
    Else
        Dim rs As New ADODB.Recordset
        rs.Open "select * from inL,SMN where txn_typ like 'RSI' and INL.INV_IDY LIKE 'RSI" & Tincrement(Text1) & "'", CON, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            Text1 = "RSI" & Tincrement(Text1)
            Command1.SetFocus
        End If
    End If
End If
End Sub

Private Sub Supplierlist()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from inL,SMN where txn_typ like 'RSI' and inl.smn_idy like smn.smn_idy ", CON, adOpenKeyset, adLockOptimistic
list1.ListItems.Clear
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("inv_idy")))
        ls.ListSubItems.Add , , Trim(rs("inv_dat"))
        ls.ListSubItems.Add , , Trim(rs("smn_nme") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Challan Not Found...!", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        Text1.SetFocus
    End If
End If
End Sub
