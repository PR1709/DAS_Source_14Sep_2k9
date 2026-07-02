VERSION 5.00
Begin VB.Form M_FRM_PRE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   960
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3435
   Icon            =   "M_FRM_PRE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   567.2
   ScaleMode       =   0  'User
   ScaleWidth      =   3225.279
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox text1 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   0
      Top             =   315
      Width           =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Prefix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   367
      Width           =   1005
   End
End
Attribute VB_Name = "M_FRM_PRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from pre", CON, adOpenStatic
If rs.RecordCount > 0 Then
If Me.Caption = "Account ID Prefix" Then
    Text1 = Trim(rs("pre_act") & "")
ElseIf Me.Caption = "Group Account ID Prefix" Then
    Text1 = Trim(rs("pre_grp") & "")
ElseIf Me.Caption = "Location ID Prefix" Then
    Text1 = Trim(rs("pre_loc") & "")
ElseIf Me.Caption = "Product ID Prefix" Then
    Text1 = Trim(rs("pre_prd") & "")
ElseIf Me.Caption = "Product Group ID Prefix" Then
    Text1 = Trim(rs("pre_cat") & "")
ElseIf Me.Caption = "Salesmen ID Prefix" Then
    Text1 = Trim(rs("pre_smn") & "")
ElseIf Me.Caption = "Zone ID Prefix" Then
    Text1 = Trim(rs("pre_zon") & "")
ElseIf Me.Caption = "Market ID Prefix" Then
    Text1 = Trim(rs("pre_mar") & "")
ElseIf Me.Caption = "Material ID Prefix" Then
    Text1 = Trim(rs("pre_mat") & "")
ElseIf Me.Caption = "Category ID Prefix" Then
    Text1 = Trim(rs("pre_cat") & "")
ElseIf Me.Caption = "Prefix Cash Sales" Then
    Text1 = Trim(rs("pre_cas") & "")
ElseIf Me.Caption = "Prefix Credit Sales" Then
    Text1 = Trim(rs("pre_cri") & "")
ElseIf Me.Caption = "Prefix Purchases" Then
    Text1 = Trim(rs("pre_pur") & "")
ElseIf Me.Caption = "Prefix Credit Note (Stocks)" Then
    Text1 = Trim(rs("pre_cns") & "")
ElseIf Me.Caption = "Prefix Debit Note (Stocks)" Then
    Text1 = Trim(rs("pre_dns") & "")
ElseIf Me.Caption = "Prefix Material Packing" Then
    Text1 = Trim(rs("pre_pkg") & "")
ElseIf Me.Caption = "Prefix Opening Stock" Then
    Text1 = Trim(rs("pre_ops") & "")
ElseIf Me.Caption = "Prefix Transfer Inwards" Then
    Text1 = Trim(rs("pre_tri") & "")
ElseIf Me.Caption = "Prefix Transfer Outwards" Then
    Text1 = Trim(rs("pre_tro") & "")
ElseIf Me.Caption = "Prefix Free Goods Issued" Then
    Text1 = Trim(rs("pre_fgi") & "")
ElseIf Me.Caption = "Prefix Free Goods Received" Then
    Text1 = Trim(rs("pre_fgr") & "")
ElseIf Me.Caption = "Prefix Orders" Then
    Text1 = Trim(rs("pre_ord") & "")
ElseIf Me.Caption = "Prefix Journal Entries" Then
    Text1 = Trim(rs("pre_jvn") & "")
ElseIf Me.Caption = "Prefix Opening Balance" Then
    Text1 = Trim(rs("pre_obn") & "")
ElseIf Me.Caption = "Prefix Bank Payments" Then
    Text1 = Trim(rs("pre_bpv") & "")
ElseIf Me.Caption = "Prefix Cash Payments" Then
    Text1 = Trim(rs("pre_cpv") & "")
ElseIf Me.Caption = "Prefix Payments" Then
    Text1 = Trim(rs("pre_pay") & "")
ElseIf Me.Caption = "Prefix Bank Receipts" Then
    Text1 = Trim(rs("pre_brv") & "")
ElseIf Me.Caption = "Prefix Cash Receipts" Then
    Text1 = Trim(rs("pre_crv") & "")
ElseIf Me.Caption = "Prefix Receipts" Then
    Text1 = Trim(rs("pre_rep") & "")
ElseIf Me.Caption = "Prefix Credit Note (Accounts)" Then
    Text1 = Trim(rs("pre_crn") & "")
ElseIf Me.Caption = "Prefix Debit Note (Accounts)" Then
    Text1 = Trim(rs("pre_drn") & "")
ElseIf Me.Caption = "Asset Type ID Prefix" Then
    Text1 = Trim(rs("pre_lev") & "")
End If
End If
End Sub

Private Sub Text1_GotFocus()
If Me.Caption = "Account ID Prefix" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Business Account Master Form"
End If
If Me.Caption = "Group Account ID Prefix" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for  Group Accounts Master Form"
End If
If Me.Caption = "Category ID Prefix" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for  Category Master Form"
End If
If Me.Caption = "Product ID Prefix" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for  Product Master Form"
End If
If Me.Caption = "Product Group ID Prefix" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for  Product Groups Master Form"
End If
If Me.Caption = "Salesmen ID Prefix" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for  Salesmen Master Form"
End If
If Me.Caption = "Zone ID Prefix" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Zone Master Form"
End If
If Me.Caption = "Market ID Prefix" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for  Market Master Form"
End If
If Me.Caption = "Category ID Prefix" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for  Category Master Form"
End If
If Me.Caption = "Prefix Cash Sales" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Cash Sales Transaction"
End If
If Me.Caption = "Prefix Credit Sales" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Credit Sales Transaction"
End If
If Me.Caption = "Prefix Purchases" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Purchases Transaction"
End If
If Me.Caption = "Prefix Credit Note (Stocks)" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Stocks Credit Note"
End If
If Me.Caption = "Prefix Debit Note (Stocks)" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Stocks Debit Note"
End If
If Me.Caption = "Prefix Material Packing" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Material Packing"
End If
If Me.Caption = "Prefix Opening Stock" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Opening Stock"
End If
If Me.Caption = "Prefix Transfer Inwards" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Transfer Inwards"
End If
If Me.Caption = "Prefix Transfer Outwards" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Transfer Outwards"
End If
If Me.Caption = "Prefix Free Goods Received" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Free Goods Received"
End If
If Me.Caption = "Prefix Free Goods Issued" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Free Goods Issued"
End If
If Me.Caption = "Prefix Orders" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Orders"
End If
If Me.Caption = "Prefix Journal Entries" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Journal Entries"
End If
If Me.Caption = "Prefix Opening Balance" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Opening Balance"
End If
If Me.Caption = "Prefix Bank Payments" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Bank Payments"
End If
If Me.Caption = "Prefix Cash Payments" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Cash Payments"
End If
If Me.Caption = "Prefix Payments" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Payments"
End If
If Me.Caption = "Prefix Bank Receipts" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Bank Receipts"
End If
If Me.Caption = "Prefix Cash Receipts" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Cash Receipts"
End If
If Me.Caption = "Prefix Receipts" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Receipts"
End If
If Me.Caption = "Prefix Credit Note (Accounts)" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Credit Note (Accounts)"
End If
If Me.Caption = "Prefix Receipts" Then
M_MDI.STS_BAR.Panels(1).Text = "Enter Prefix for Accounts Debit Note"
End If
End Sub

Private Sub TEXT1_KeyDown(K As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from pre", CON, adOpenKeyset, adLockPessimistic
        If rs.RecordCount = 0 Then rs.AddNew
        If Me.Caption = "Account ID Prefix" Then
            rs("pre_act") = Trim(Text1)
        ElseIf Me.Caption = "Group Account ID Prefix" Then
            rs("pre_grp") = Trim(Text1)
        ElseIf Me.Caption = "Location ID Prefix" Then
            rs("pre_loc") = Trim(Text1)
        ElseIf Me.Caption = "Product ID Prefix" Then
            rs("pre_prd") = Trim(Text1)
        ElseIf Me.Caption = "Product Group ID Prefix" Then
            rs("pre_cat") = Trim(Text1)
        ElseIf Me.Caption = "Salesmen ID Prefix" Then
            rs("pre_smn") = Trim(Text1)
        ElseIf Me.Caption = "Zone ID Prefix" Then
            rs("pre_zon") = Trim(Text1)
        ElseIf Me.Caption = "Market ID Prefix" Then
            rs("pre_mar") = Trim(Text1)
        ElseIf Me.Caption = "Category ID Prefix" Then
            rs("pre_cat") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Cash Sales" Then
            rs("pre_cas") = Trim(Text1)
        ElseIf Me.Caption = "Material ID Prefix" Then
            rs("pre_mat") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Credit Sales" Then
            rs("pre_cri") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Purchases" Then
            rs("pre_pur") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Credit Note (Stocks)" Then
            rs("pre_cns") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Debit Note (Stocks)" Then
            rs("pre_dns") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Material Packing" Then
            rs("pre_pkg") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Opening Stock" Then
            rs("pre_ops") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Transfer Inwards" Then
            rs("pre_tri") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Transfer Outwards" Then
            rs("pre_tro") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Free Goods Issued" Then
            rs("pre_fgi") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Free Goods Received" Then
            rs("pre_fgr") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Orders" Then
            rs("pre_ord") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Journal Entries" Then
            rs("pre_jvn") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Opening Balance" Then
            rs("pre_obn") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Bank Payments" Then
            rs("pre_bpv") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Cash Payments" Then
            rs("pre_cpv") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Payments" Then
            rs("pre_pay") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Bank Receipts" Then
            rs("pre_brv") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Cash Receipts" Then
            rs("pre_crv") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Receipts" Then
            rs("pre_rep") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Credit Note (Accounts)" Then
            rs("pre_crn") = Trim(Text1)
        ElseIf Me.Caption = "Prefix Debit Note (Accounts)" Then
            rs("pre_drn") = Trim(Text1)
        ElseIf Me.Caption = "Asset Type ID Prefix" Then
            rs("PRE_LEV") = Trim(Text1)
        End If
        rs.Update
        Unload Me
ElseIf K = 27 Then
    Unload Me
    Set Form = Nothing
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
If KeyAscii = 39 Or KeyAscii = 41 Then KeyAscii = 0
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii > 64 And KeyAscii < 92) Or KeyAscii = 8 Then
Else
KeyAscii = 0
End If
End Sub
