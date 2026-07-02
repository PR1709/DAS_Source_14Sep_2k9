VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_REP_ASM 
   Caption         =   "Select Asset"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "M_REP_ASM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Select Asset"
      TabPicture(0)   =   "M_REP_ASM.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_REP_ASM.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "Asset Details"
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   3135
         Begin VB.TextBox ASTIDY 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   247
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   690
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Select Asset"
            Height          =   195
            Left            =   165
            TabIndex        =   10
            Top             =   315
            Width           =   885
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Asset Code"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   750
            Width           =   810
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   3360
         TabIndex        =   2
         Top             =   600
         Width           =   1215
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView List1 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   1
         Top             =   840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2143
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "Order ID"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order Date"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Asset"
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
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   4215
      End
   End
End
Attribute VB_Name = "M_REP_ASM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ASTNME, ASTCODE

Private Sub Command1_Click()
Command1.Enabled = False
con.Execute "UPDATE DIR SET DIR_MSG='" & Trim(ASTNME) & "' WHERE DIR_IDY LIKE 'DLR' AND DIR_IDX LIKE 'IDY'", d
con.Execute "UPDATE DIR SET DIR_MSG='" & Trim(Combo1) & "' WHERE DIR_IDY LIKE 'LOC' AND DIR_IDX LIKE 'IDY'", d
OpenReport CURDIR & "\Assets.RP1", "Asset Tracking", 2, "AST->ASM_IDY = '" & Trim(ASTIDY) & "'.AND.ast->asm_cod='" & Trim(Combo1) & "'", 6
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
ASTIDY.SetFocus
End Sub

Private Sub Form_KeyDown(KeyAscii As Integer, s As Integer)
If KeyAscii = 27 Then
    If Me.ActiveControl.name = "Listview1" Then
        SSTab1.Tab = 0
'        Text1.SetFocus
    Else
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
End Sub

Private Sub List1_DblClick()
List1_keypress 13
End Sub

Private Sub List1_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list1.ListItems.count > 0 Then
        ASTIDY = list1.SelectedItem.ListSubItems.Item(1)
        ASTNME = list1.SelectedItem
        Dim REC_SET As New ADODB.Recordset
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT ASM_COD FROM AST WHERE ASM_IDY LIKE '" & Trim(ASTIDY) & "' GROUP BY ASM_COD", con, adOpenKeyset, adLockPessimistic
        Combo1.CLEAR
        For A = 1 To REC_SET.RecordCount
            Combo1.AddItem Trim(REC_SET(0) & "")
            REC_SET.MoveNext
        Next
        SSTab1.Tab = 0
        Combo1.SetFocus
    End If
End If
End Sub

Private Sub ASTIDY_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub ASM_LST()
If CHECKTB("asm.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from ASM where ASM_nme like '" & Trim(Text1) & "%' AND IND_IDY='Y'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ColumnHeaders.CLEAR
    list1.ColumnHeaders.Add , , "Asset Name"
    list1.ColumnHeaders.Add , , "ID"
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("ASM_nme")))
        ls.ListSubItems.Add , , Trim(rs("ASM_idy"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "Asset Not Found, Define Asset !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    ASTIDY.SetFocus
    
'    Text1.Enabled = True
'    Text1.SetFocus
End If
End Sub

Private Sub ASTIDY_KeyDown(K As Integer, Shift As Integer)
Dim REC_SET As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("asm.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        
        Exit Sub
    End If

        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from ASM where ASM_idy like '" & Sincrement(Trim(ASTIDY)) & "' AND IND_IDY='Y'", con, adOpenKeyset, adLockPessimistic
        If REC_SET.RecordCount > 0 Then
            ASTIDY = Sincrement(Trim(ASTIDY))
            ASTNME = Trim(REC_SET!ASM_NME)
            If REC_SET.State = 1 Then REC_SET.Close
            REC_SET.Open "SELECT ASM_COD FROM AST WHERE ASM_IDY LIKE '" & Trim(ASTIDY) & "' GROUP BY ASM_COD", con, adOpenKeyset, adLockPessimistic
            Combo1.CLEAR
            For A = 1 To REC_SET.RecordCount
                Combo1.AddItem Trim(REC_SET(0) & "")
                REC_SET.MoveNext
            Next
        Else
            ASM_LST
        End If
End If
End Sub

Private Sub ASTIDY_KEYPRESS(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

