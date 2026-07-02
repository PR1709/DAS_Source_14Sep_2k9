VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form M_FRM_SSI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empties Shell Sorting"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "M_FRM_SSI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   7560
      TabIndex        =   9
      Top             =   4595
      Width           =   1190
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   930
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Print"
         Height          =   425
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   930
      End
   End
   Begin MSFlexGridLib.MSFlexGrid ms1 
      Height          =   1455
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   7
      BackColor       =   -2147483633
   End
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MS 
      Height          =   5115
      Left            =   285
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   9022
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      RowHeightMin    =   330
      BackColor       =   -2147483633
      BackColorSel    =   14737632
      ForeColorSel    =   8388608
      GridColor       =   8421504
      GridColorFixed  =   4210752
      FocusRect       =   2
      HighLight       =   0
      GridLines       =   2
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
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   7560
      TabIndex        =   2
      Top             =   3780
      Width           =   1190
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   400
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7680
      TabIndex        =   8
      Top             =   4560
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Packs"
      Height          =   195
      Left            =   6600
      TabIndex        =   7
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5760
      TabIndex        =   6
      Top             =   4560
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Shells"
      Height          =   195
      Left            =   4680
      TabIndex        =   5
      Top             =   4560
      Width           =   825
   End
End
Attribute VB_Name = "M_FRM_SSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
If MsgBox("Confirm Glass Sort Save?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
    Chkshell = chkShells
    If Chkshell = False Then
        MsgBox "Invalid Glass Sort!", vbExclamation, "DAS Version 4.0"
        MS.SetFocus
        Exit Sub
    End If
    con.Execute "delete from ssi where inv_idy like '" & Trim(MS.TextMatrix(1, 5)) & "'"
    rs.Open "select * from ssi", con, adOpenKeyset, adLockPessimistic
    
    invdat = GetValue("inv_dat", "inl where inv_idy like '" & Trim(MS.TextMatrix(1, 5)) & "'")
    If IsDate(invdat) = False Then
        invdat = GetValue("inv_dat", "SNV where TXN_idy like '" & Trim(MS.TextMatrix(1, 5)) & "'")
    End If
    For A = 1 To MS.Rows - 1
'        If Val(MS.TextMatrix(A, 3)) > 0 Or Val(MS.TextMatrix(A, 4)) > 0 Then
            rs.AddNew
            rs!INV_IDY = MS.TextMatrix(A, 5)
            rs!prd_idy = MS.TextMatrix(A, 10)
            rs!shl_nos = Val(MS.TextMatrix(A, 3))
            rs!pck_nos = Val(MS.TextMatrix(A, 4))
            rs!PRD_PCK = MS.TextMatrix(A, 1)
            rs!SUB_UNT = MS.TextMatrix(A, 6)
            rs!INV_DAT = invdat
            rs!TXN_TYP = MS.TextMatrix(A, 7)
            rs!shl_idy = MS.TextMatrix(A, 9)
            rs!pck_pcs = MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 6)))
            rs.Update
'        End If
    Next
    
    'Unload M_FRM_TRS
'    Command3_Click
    Dim SQRY As String
    SQRY = "SSI->INV_IDY = '" & Trim(MS.TextMatrix(1, 5)) & "' .AND. (SSI->TXN_TYP = 'AOU')"
    OpenReport CURDIR & "\SNP.RP1", "Shell Sort Transaction", 2, SQRY, 6
    Unload Me
    'Unload M_FRM_TRS
Else
    Unload Me
End If
End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim SQRY As String
'sqry = "ssi->inv_idy='" & Trim(M_FRM_TRS.invid) & "'.AND.(SSI->txn_typ='PUR'.OR.SSI->TXN_TYP='STI'.OR.SSI->TXN_TYP='SRN'.OR.SSI->TXN_TYP='DPI'.OR.SSI->TXN_TYP='AOU')"
SQRY = "SSI->INV_IDY = '" & Trim(M_FRM_TRS.invid) & "'.AND.(SSI->TXN_TYP = 'AOU')"
OpenReport CURDIR & "\SNP.RP1", "Shell Sort Transaction", 2, SQRY, 6
End Sub

Private Sub Form_Activate()
MS.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Command1_Click
End Sub

Private Sub Form_Load()
With MS
.FormatString = "<Shell Name|Pack |<Product Name|>Shells|>Packs"
MS.MergeCells = flexMergeRestrictColumns
MS.MergeCol(0) = True
MS.MergeCol(1) = True
MS.FixedCols = 2
.ColWidth(0) = 1510
.ColWidth(1) = 1220
.ColAlignment(1) = 1
.ColWidth(2) = 2600
.ColWidth(3) = 900
.ColWidth(4) = 960
.ColWidth(5) = 0
.ColWidth(6) = 0
.ColWidth(7) = 0
.ColWidth(8) = 0
.ColWidth(9) = 0
.ColWidth(10) = 0
End With
LoadSKU
End Sub

Public Sub LoadSKU()
Dim rs As New ADODB.Recordset
With MS
STR1 = Split(Gltxninvid, ",")
Gltxninvid = ""
'RS.Open "select * from prd,shl where shl.prd_idy =prd.prd_idy", CON, adOpenKeyset, adlockpessimistic
rs.Open "select * from prd,shl where shl.prd_idy =prd.prd_idy", con, adOpenKeyset, adLockPessimistic
.Rows = rs.RecordCount + 1
For A = 1 To rs.RecordCount
    .TextMatrix(A, 0) = Trim(rs!shl_nme)
    .TextMatrix(A, 1) = Trim(rs!PRD_PCK)
    .TextMatrix(A, 2) = Trim(rs!PRD_NME)
    .TextMatrix(A, 3) = "" 'Val(RS!shl_nos)
    .TextMatrix(A, 4) = "" 'QTY(RS!pck_nos)
    .TextMatrix(A, 5) = Trim(STR1(1))
    .TextMatrix(A, 6) = Val(rs!SUB_UNT)
    .TextMatrix(A, 7) = STR1(0)
    .TextMatrix(A, 9) = Trim(rs!shl_idy)
    .TextMatrix(A, 10) = Trim(rs!prd_idy)
    rs.MoveNext
Next
End With
If rs.State = 1 Then rs.Close
rs.Open "select * from ssi where inv_idy = '" & Trim(STR1(1)) & "'", con, adOpenStatic, adLockPessimistic
Do While rs.EOF = False
    For A = 1 To MS.Rows - 1
        If MS.TextMatrix(A, 10) = rs!prd_idy Then
            MS.TextMatrix(A, 3) = Val(rs!shl_nos)
            MS.TextMatrix(A, 4) = QTY(rs!pck_nos)
        End If
    Next
    rs.MoveNext
Loop
Set rs = Nothing
MS1.Rows = 0
If rs.State = 1 Then rs.Close
If STR1(0) = "DPI" Or STR1(0) = "AOU" Then
    rs.Open "select distinct SNV.PRD_PCK,SNV.PRD_UPC,SUM(R_SHL_NOS),SUM(R_PCK_NOS),SNV.SHL_IDY,SUM(DP_SHL_NOS),SUM(DP_PCK_NOS),SUM(D_SHL_NOS),SUM(D_PCK_NOS) from snv where txn_idy like '" & Trim(STR1(1)) & "' and txn_typ like '" & Trim(STR1(0)) & "' Group by snv.prd_pck,SNV.PRD_UPC,snv.SHL_IDY", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "select distinct SNV.PRD_PCK,SNV.PRD_UPC,SUM(R_SHL_NOS),SUM(R_PCK_NOS),SNV.SHL_IDY,SUM(DP_SHL_NOS),SUM(DP_PCK_NOS) from snv where inv_idy like '" & Trim(STR1(1)) & "' and txn_typ like '" & Trim(STR1(0)) & "' Group by snv.prd_pck,SNV.PRD_UPC,snv.SHL_IDY", con, adOpenKeyset, adLockPessimistic
End If
MS1.Rows = rs.RecordCount + 1
For A = 1 To rs.RecordCount
    MS1.TextMatrix(A, 0) = Trim(rs(0))
    MS1.TextMatrix(A, 1) = Trim(rs(1))
    MS1.TextMatrix(A, 2) = Trim(rs(4))
    If STR1(0) = "DPI" Then
        MS1.TextMatrix(A, 3) = Trim(rs(5) & "")
        MS1.TextMatrix(A, 4) = Val(rs(6) & "")
    ElseIf STR1(0) = "AOU" Then
        MS1.TextMatrix(A, 3) = Trim(rs(7) & "")
        MS1.TextMatrix(A, 4) = Val(rs(8) & "")
    Else
        MS1.TextMatrix(A, 3) = Trim(rs(2) & "")
        MS1.TextMatrix(A, 4) = Val(rs(3) & "")
    End If
    rs.MoveNext
Next
Total
End Sub

Private Sub MS_KeyPress(K As Integer)
If (MS.Col = 3 Or MS.Col = 4) And MS.Row > 0 Then
    If K > 47 And K < 58 Or K = 46 Then
        TXT = ""
        TXT.MaxLength = 9
        TXT.Visible = True
        TXT.Top = MS.Top + MS.CellTop
        TXT.Left = MS.CellLeft + MS.Left
        TXT.Width = MS.CellWidth
        TXT.Height = MS.CellHeight
        TXT.Visible = True
        TXT.SetFocus
        If K = vbKeySubtract Then
        Else
            TXT = Chr(K)
            TXT.SelStart = Len(TXT)
        End If
    End If
End If
End Sub
Private Sub MS_LeaveCell()
TXT.Visible = False
End Sub

Private Sub Txt_Keypress(K As Integer)
If K = 13 Then
    If MS.Col = 3 Then
        MS.Text = Val(TXT)
        MS.TextMatrix(MS.Row, 4) = QTY(Val(TXT))
        MS.Col = 4
        TXT.Visible = False
    Else
        If CHKSUBUNT(Val(TXT), Val(MS.TextMatrix(MS.Row, 6))) = True Then
            MS.Text = QTY(TXT)
            If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
            MS.Col = 3
            
            TXT.Visible = False
        Else
            SendKeys "{home}+{end}"
            TXT.SetFocus
        End If
    End If
    Total
    MS.SetFocus
End If
End Sub
Public Function chkShells() As Boolean
For b = 1 To MS1.Rows - 1
    MS1.TextMatrix(b, 5) = 0
    MS1.TextMatrix(b, 6) = 0
Next
For A = 1 To MS.Rows - 1
    For b = 1 To MS1.Rows - 1
        If Trim(MS.TextMatrix(A, 1)) = Trim(MS1.TextMatrix(b, 0)) And Val(MS.TextMatrix(A, 6)) = Val(MS1.TextMatrix(b, 1)) And Trim(MS.TextMatrix(A, 9)) = Trim(MS1.TextMatrix(b, 2)) Then
            MS1.TextMatrix(b, 5) = Val(MS1.TextMatrix(b, 5)) + Val(MS.TextMatrix(A, 3))
            MS1.TextMatrix(b, 6) = Val(MS1.TextMatrix(b, 6)) + MTS(Val(MS.TextMatrix(A, 4)), Val(MS.TextMatrix(A, 6)))
        End If
    Next
Next
    For b = 1 To MS1.Rows - 1
        A3 = Val(A3) + Val(MS1.TextMatrix(b, 3))
        A4 = Val(A4) + Val(MS1.TextMatrix(b, 4))
        A5 = Val(A5) + Val(MS1.TextMatrix(b, 5))
        A6 = Val(A6) + Val(MS1.TextMatrix(b, 6))
    Next
    If Val(A3) <> Label2 Then
        chkShells = False
        Exit Function
    End If
    chkShells = True
End Function

Public Sub Total()
Label2 = 0
Label4 = 0
For b = 1 To MS.Rows - 1
    Label2 = Val(Label2) + Val(MS.TextMatrix(b, 3))
    Label4 = Val(Label4) + Val(MS.TextMatrix(b, 4))
Next
Label4 = QTY(Label4)
End Sub
