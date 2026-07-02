VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_ORD 
   Caption         =   "Supplier Order Request"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   Icon            =   "M_FRM_ORD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Order Details"
      TabPicture(0)   =   "M_FRM_ORD.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SUP_LBL"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ORD_DAT"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "MS"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "M_FLX_AMT"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "SUP_IDY"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ORD_IDY"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_ORD.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ORD_LST"
      Tab(1).Control(1)=   "SUP_LST"
      Tab(1).Control(2)=   "LBL_NME"
      Tab(1).ControlCount=   3
      Begin MSComctlLib.ListView ORD_LST 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   6376
         View            =   3
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
            Text            =   "Order ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox ORD_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1560
         TabIndex        =   11
         Top             =   622
         Width           =   1095
      End
      Begin VB.TextBox SUP_IDY 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox M_FLX_AMT 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   4320
         TabIndex        =   1
         Top             =   2760
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   2415
         Left            =   630
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1680
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   2
         ScrollBars      =   2
         AllowUserResizing=   1
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
      Begin MSComctlLib.ListView SUP_LST 
         Height          =   3600
         Left            =   -74760
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   6350
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
            Text            =   "Supplier Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3175
         EndProperty
      End
      Begin MSComCtl2.DTPicker ORD_DAT 
         Height          =   330
         Left            =   4440
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   615
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16580611
         CurrentDate     =   37258
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Order ID"
         Height          =   195
         Left            =   810
         TabIndex        =   10
         Top             =   690
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Left            =   3960
         TabIndex        =   9
         Top             =   683
         Width           =   345
      End
      Begin VB.Label SUP_LBL 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3960
         TabIndex        =   8
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplier ID"
         Height          =   195
         Left            =   630
         TabIndex        =   7
         Top             =   1148
         Width           =   780
      End
      Begin VB.Label LBL_NME 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -74760
         TabIndex        =   6
         Top             =   600
         Width           =   7905
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   120
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
            Picture         =   "M_FRM_ORD.frx":047A
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ORD.frx":09BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ORD.frx":0D86
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ORD.frx":10DA
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ORD.frx":1712
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ORD.frx":1A66
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ORD.frx":1EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ORD.frx":21BA
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ORD.frx":25C6
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ORD.frx":26D2
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ORD.frx":2A26
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_ORD.frx":2E3A
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8355
      _ExtentX        =   14737
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
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Category"
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
         Caption         =   "&Qty -> Order"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7000
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   1335
      End
   End
End
Attribute VB_Name = "M_FRM_ORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD_REC As Boolean

Private Sub Command1_Click()
Command1.Enabled = False
For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 3)) = 0 Then
        MS.TextMatrix(A, 3) = QTY(Val(MS.TextMatrix(A, 2)))
    End If
Next
MS.SetFocus
End Sub

Private Sub MS_GotFocus()
'MS.HighLight = flexHighlightAlways
End Sub

Private Sub MS_KeyPress(K As Integer)
If MS.Col = 3 And MS.Row > 0 Then
    If K > 47 And K < 58 Or K = 46 Then
        M_FLX_AMT = "0.000"
        M_FLX_AMT.MaxLength = 9
        M_FLX_AMT.Visible = True
        M_FLX_AMT.Top = MS.Top + MS.CellTop
        M_FLX_AMT.Left = MS.CellLeft + MS.Left
        M_FLX_AMT.Width = MS.CellWidth
        M_FLX_AMT.Height = MS.CellHeight
        M_FLX_AMT.Visible = True
        M_FLX_AMT.SetFocus
        M_FLX_AMT = Chr(K)
        M_FLX_AMT.SelStart = Len(M_FLX_AMT)
    End If
End If
End Sub

Private Sub MS_LeaveCell()
If M_FLX_AMT.Visible = True Then M_FLX_AMT.Visible = False
End Sub

Public Sub DISP_PRD()
If MS.Rows < 3 Then
If CHECKTB("act.dbf,prd.dbf,cld.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim REC_SET_P As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT ACT_NME,ACT_IDY FROM ACT WHERE ACT_IDY LIKE '" & Trim(SUP_IDY) & "' AND GRP_CLF LIKE 'Creditor' AND DBT_TYP LIKE 'Y'", con, adOpenKeyset, adLockPessimistic
If REC_SET_P.State = 1 Then REC_SET_P.Close
REC_SET_P.Open "SELECT * FROM PRD WHERE CMP_IDY LIKE '" & Trim(REC_SET("ACT_IDY")) & "' AND STK_TYP LIKE 'Salable'", con, adOpenKeyset, adLockPessimistic
For A = 1 To REC_SET_P.RecordCount
    MS.AddItem ""
    MS.TextMatrix(A, 0) = Trim(REC_SET_P("PRD_IDY"))
    MS.TextMatrix(A, 1) = REC_SET_P("PRD_NME")
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT CLS_STK FROM CLD WHERE PRD_IDY LIKE '" & Trim(REC_SET_P("PRD_IDY")) & "'", con, adOpenDynamic, adLockPessimistic
    MS.TextMatrix(A, 2) = QTY(REC_SET_P("CUR_LEV"))
    If IsNull(Trim(REC_SET_P("REO_LEV"))) = False Then          'Added by PNR for below commented code - Invalid NULL error
        MS.TextMatrix(A, 3) = QTY(Trim(REC_SET_P("REO_LEV")))
    Else
        MS.TextMatrix(A, 3) = "0.000"
    End If
    MS.TextMatrix(A, 4) = REC_SET_P("SUB_UNT")
    REC_SET_P.MoveNext
Next
End If
MS.SetFocus
End Sub

Public Sub DISP_ORD()
If MS.Rows < 3 Then
If CHECKTB("ord.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim REC_SET_O As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
If Len(Trim(ORD_IDY)) = 10 Then
    REC_SET.Open "SELECT * FROM ORD WHERE ORD_IDY LIKE '" & Trim(ORD_IDY) & "'", con, adOpenKeyset, adLockPessimistic
Else
    REC_SET.Open "SELECT * FROM ORD WHERE ORD_IDY LIKE 'ORD" & Tincrement(ORD_IDY) & "'", con, adOpenKeyset, adLockPessimistic
End If
    If REC_SET.RecordCount > 0 Then
    If CHECKTB("prd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        Exit Sub
    End If
    ORD_IDY = Trim(REC_SET("ORD_IDY"))
    ORD_IDY.Enabled = False
    ORD_DAT = Trim(REC_SET("ORD_DAT"))
    SUP_IDY = Trim(REC_SET("ACT_IDY"))
    MS.Rows = REC_SET.RecordCount + 1
        For A = 1 To MS.Rows - 1
            If REC_SET_O.State = 1 Then REC_SET_O.Close
            REC_SET_O.Open "SELECT * FROM PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET("PRD_IDY")) & "' and stk_typ like 'Salable'", con, adOpenKeyset, adLockPessimistic
            MS.AddItem ""
            MS.TextMatrix(A, 0) = Trim(REC_SET("PRD_IDY"))
            MS.TextMatrix(A, 1) = REC_SET_O("PRD_NME")
            MS.TextMatrix(A, 2) = QTY(REC_SET_O("CUR_LEV"))
            MS.TextMatrix(A, 3) = QTY(Trim(REC_SET("ORD_STK")))
            REC_SET.MoveNext
        Next
    End If
MS.Rows = REC_SET.RecordCount + 1
End If
MS.SetFocus
End Sub

Private Sub M_FLX_AMT_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    If Val(CMP()) >= Val(MS.TextMatrix(MS.Row, 4)) Then
        MsgBox "Sub Units are " & Val(MS.TextMatrix(MS.Row, 4)), vbExclamation, "DAS Version 4.0"
        M_FLX_AMT.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    Else
        MS.Text = QTY(M_FLX_AMT)
        M_FLX_AMT.Visible = False
    End If
    If MS.Row = MS.Rows - 1 Then
    Else
        MS.Row = MS.Row + 1
    End If
    MS.SetFocus
End If
End Sub

Private Sub M_FLX_AMT_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 46 Or K = 13 Then
Else
K = 0
End If
End Sub

Private Sub Form_Activate()
checkdongle
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
ORD_LST.Visible = False
SUP_LST.Visible = False
ORD_IDY.Enabled = True
ORD_IDY.SetFocus
MS.ColWidth(4) = 0
MS.Rows = 1
End Sub

Private Sub V_SUP_LST()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from act where ACT_NME like '" & Trim(SUP_IDY) & "%' AND GRP_CLF LIKE 'Creditor' AND DBT_TYP LIKE 'Y'", con, adOpenKeyset, adLockPessimistic
SUP_LST.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = SUP_LST.ListItems.Add(, , Trim(REC_SET("act_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("act_idy"))
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    SUP_LST.Visible = True
    LBL_NME.Caption = "Select Supplier"
    SUP_LST.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Supplier Not Found, Define Supplier !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        SUP_IDY.SetFocus
    End If
End If
End Sub

Private Sub V_ORD_LST()
If CHECKTB("ord.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
Dim ORDDAT1, ORDDAT2 As String
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select DISTINCT ORD_IDY,ORD_DAT from ORD where ORD_IDY like '" & Trim(ORD_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
ORD_LST.ListItems.CLEAR
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        ORDDAT1 = Format(REC_SET("ORD_DAT"), "MM/DD/YYYY")
        ORDDAT2 = Format(ORDDAT1, "DD/MM/YYYY")
        Set ls = ORD_LST.ListItems.Add(, , Trim(REC_SET("ORD_IDY")))
        ls.ListSubItems.Add , , Trim(ORDDAT2)
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    ORD_LST.Visible = True
    LBL_NME.Caption = "Select Order"
    ORD_LST.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Order Not Found, Define Order !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        ORD_IDY.SetFocus
    End If
End If
End Sub

Public Sub Printrecord()
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_ORD") = True Then
    M_FRM_PRV.Check3.Value = 1
Else
    M_FRM_PRV.Check3.Value = 0
End If
M_FRM_PRV.Tag = "ORD"
M_FRM_PRV.Show 1
End Sub

Private Sub Form_KeyDown(K As Integer, Shift As Integer)
If K = 17 Then K = 0
If K = 27 Then
    If SSTab1.Tab = 1 Then
        SSTab1.Tab = 0
        ORD_IDY.Enabled = True
        ORD_IDY.SetFocus
    Else
        If M_FRM_ORD.ActiveControl.name = "ORD_IDY" Then
            Unload Me
        Else
            saverecord
        End If
    End If
    Exit Sub
End If
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If ORD_LST.Visible = False And SUP_LST.Visible = False Then
    If K = 78 And Shift = 2 Then newrec
    If K = 80 And Shift = 2 Then Printrecord
    If K = 83 And Shift = 2 Then saverecord
    If K = 88 And Shift = 2 Then Unload Me
    If K = 81 And Shift = 2 Then CANCELRECORD
    If K = 112 Then M_HELP
    If K = 69 And Shift = 2 Then
        If M_FRM_ORD.ActiveControl.name = "ORD_IDY" Then
            V_ORD_LST
        ElseIf M_FRM_ORD.ActiveControl.name = "SUP_IDY" Then
            V_SUP_LST
        End If
    End If
End If
If K > 111 And K < 120 Then If TypeOf M_FRM_ORD.ActiveControl Is TextBox Then M_FRM_ORD.ActiveControl.Text = FUNKEY(K)
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If SSTab1.Tab = 1 Then
        SSTab1.Tab = 0
        ORD_IDY.Enabled = True
        ORD_IDY.SetFocus
    Else
        If M_FRM_ORD.ActiveControl.name = "ORD_IDY" Then
            Unload Me
        Else
            saverecord
        End If
    End If
End If
End Sub

Private Sub Form_Load()
MS.TextMatrix(0, 0) = "Prd - SKU ID"
MS.ColAlignment(0) = 1
MS.TextMatrix(0, 1) = "Product - SKU Name"
MS.TextMatrix(0, 2) = "Qty"
MS.TextMatrix(0, 3) = "Order"
SCl MS, 3
MS.ColAlignment(2) = 7
MS.ColAlignment(3) = 7
MS.ColWidth(0) = 1200
MS.ColWidth(1) = 2650
MS.ColWidth(2) = 1500
MS.ColWidth(3) = 1350
M_FLX_AMT = "0.000"
ORD_DAT = Date
MOD_REC = False
End Sub

Private Sub MS_Scroll()
If M_FLX_AMT.Visible = True Then
    M_FLX_AMT = "0.000"
    M_FLX_AMT.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub ORD_IDY_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("ord.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        Exit Sub
    End If
    Dim ls As ListItem
    Dim rs As New ADODB.Recordset
    rs.Open "select DISTINCT ORD_IDY,ORD_DAT from ORD where ORD_IDY like 'ORD" & Tincrement(ORD_IDY) & "%'", con, adOpenStatic
    ORD_LST.ListItems.CLEAR
    If rs.RecordCount > 0 Then
        DISP_ORD
    Else
        V_ORD_LST
    End If
    MOD_REC = True
End If
End Sub

Private Sub ORD_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub SUP_LST_DblClick()
SUP_LST_KeyPress 13
End Sub

Private Sub SUP_LST_GotFocus()
SUP_LST.Visible = True
SUP_LST.SetFocus
End Sub

Private Sub SUP_LST_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If SUP_LST.ListItems.count > 0 Then
    SUP_IDY = SUP_LST.SelectedItem.ListSubItems(1)
    SUP_IDY_KeyDown 13, 0
    SUP_LST.Visible = False
    SSTab1.Tab = 0
    DISP_PRD
    MS.Col = 3
    MS.SetFocus
    End If
End If
End Sub

Private Sub ord_LST_DblClick()
ord_LST_KeyPress 13
End Sub

Private Sub ord_LST_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ORD_LST.ListItems.count > 0 Then
    ORD_IDY = ORD_LST.SelectedItem
    ORD_IDY.Enabled = False
    ORD_LST.Visible = False
    SSTab1.Tab = 0
    DISP_ORD
    MS.Col = 3
    MS.SetFocus
    End If
End If
End Sub

Private Sub saverecord()
Dim rs As New ADODB.Recordset
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If MsgBox("Confirm Transaction Save ? ", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    CANCELRECORD
    Exit Sub
End If
If CHECKTB("ord.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
If M_FLX_AMT.Visible = True Then
    MS.Text = M_FLX_AMT
    M_FLX_AMT.Visible = False
    M_FLX_AMT = "0.000"
End If
If Len(Trim(SUP_IDY)) = 0 Then
    MsgBox "Select Supplier ID", vbExclamation, "DAS Version 4.0"
    SUP_IDY.Enabled = True
    SUP_IDY.SetFocus
    Exit Sub
End If
If MOD_REC = False Then
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM ORD", con, adOpenKeyset, adLockPessimistic
    INC = "ORD" & T7increment("IDY_OTC")
    For A = 1 To MS.Rows - 1
        If Trim(MS.TextMatrix(A, 0)) <> "" Then
            If Trim(MS.TextMatrix(A, 3)) <> 0 Then
                rs.AddNew
                rs("ORD_IDY") = INC
                rs("ORD_DAT") = Trim(ORD_DAT)
                rs("ACT_IDY") = Trim(SUP_IDY)
                rs("PRD_IDY") = Trim(MS.TextMatrix(A, 0))
                rs("CLS_STK") = Trim(MS.TextMatrix(A, 2))
                rs("ORD_STK") = Trim(MS.TextMatrix(A, 3))
            End If
            rs.Update
        End If
    Next
    MsgBox "New Order ID is " & rs("ord_idy"), vbExclamation, "DAS Version 4.0"
Else
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM ORD WHERE ORD_IDY LIKE '" & Trim(ORD_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        If CHECKTB("prd.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            Exit Sub
        End If
    MS.Rows = rs.RecordCount + 1
    For A = 1 To MS.Rows - 1
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM PRD WHERE PRD_IDY LIKE '" & Trim(rs("PRD_IDY")) & "' and stk_typ like 'Salable'", con, adOpenKeyset, adLockPessimistic
        If Trim(MS.TextMatrix(A, 0)) <> "" Then
        INC = Trim(ORD_IDY)
        rs("ORD_DAT") = Trim(ORD_DAT)
        rs("PRD_IDY") = Trim(MS.TextMatrix(A, 0))
        rs("CLS_STK") = Trim(MS.TextMatrix(A, 2))
        rs("ORD_STK") = Trim(MS.TextMatrix(A, 3))
        rs("RCT_QTY") = 0#
        rs.Update
        rs.MoveNext
        End If
    Next
    End If
End If
If CHECKTB("prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
If check("PRT", "PRT_NME", "M_FRM_ORD") = True Then
   OpenReport CURDIR & "\REP_LIB.RP1", "Supplier Order", 2, "ORD->ORD_IDY='" & Trim(INC) & "'", 6
End If
clrctr M_FRM_ORD
setval
CANCELRECORD
SSTab1.Tab = 0
ORD_IDY.Enabled = True
ORD_IDY.SetFocus
MOD_REC = False
End Sub

Private Sub CANCELRECORD()
MOD_REC = False
clrctr M_FRM_ORD
setval
Form_Load
MS.Rows = 1
SSTab1.Tab = 0
ORD_DAT = Date
ORD_IDY.Enabled = True
ORD_IDY.SetFocus
Command1.Enabled = False
End Sub

Public Sub setval()
clrctr M_FRM_ORD
SUP_IDY = ""
ORD_IDY = ""
End Sub

Private Sub ORD_DAT_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    SUP_IDY.Enabled = True
    SUP_IDY.SetFocus
End If
End Sub

Private Sub SUP_IDY_Change()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim REC_SUP As New ADODB.Recordset
If REC_SUP.State = 1 Then REC_SUP.Close
If SUP_IDY.Text = "" Then
SUP_LBL.Caption = ""
Else
REC_SUP.Open "SELECT ACT_NME FROM ACT WHERE ACT_IDY LIKE '" & Trim(SUP_IDY) & "'", con, adOpenDynamic, adLockPessimistic
If Not REC_SUP.EOF Then SUP_LBL.Caption = Trim(REC_SUP("ACT_NME")) & ""
End If
End Sub

Private Sub SUP_IDY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 27 Then
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
rs.Open "select * from ACT where ACT_IDY like '" & Trim(SUP_IDY) & "%' AND GRP_CLF LIKE 'Creditor' AND DBT_TYP LIKE 'Y'", con, adOpenStatic
SUP_LST.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = SUP_LST.ListItems.Add(, , Trim(rs("ACT_NME")))
        ls.ListSubItems.Add , , Trim(rs("ACT_IDY") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    SUP_LST.ColumnHeaders(1).Text = "Supplier Name"
    SUP_LST.ColumnHeaders(2).Text = "Supplier ID"
    LBL_NME.Caption = "Select Supplier"
    ORD_LST.Visible = False
    SUP_LST.Visible = True
    SUP_LST.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Supplier Not Found, Define Supplier !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        SUP_IDY.Enabled = True
        SUP_IDY.SetFocus
    End If
End If
End If
End Sub

Private Sub SUP_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 13 Then ORD_DAT.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    newrec
ElseIf Button.KEY = "s" Then
    saverecord
ElseIf Button.KEY = "c" Then
    CANCELRECORD
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If M_FRM_ORD.ActiveControl.name = "ORD_IDY" Then ORD_IDY_KeyDown 13, 0
ElseIf Button.KEY = "p" Then
    Printrecord
ElseIf Button.KEY = "pr" Then
    Exit Sub
er1:
    MsgBox "Order Not Found, Create Order Transaction!", vbExclamation, "DAS Version 4.0"
'ElseIf Button.KEY = "crt" Then
'    CREATE
End If
End Sub

Public Sub newrec()
CANCELRECORD
ORD_IDY.Enabled = False
ORD_DAT.SetFocus
SSTab1.Tab = 0
MOD_REC = False
Command1.Enabled = True
End Sub
Public Function CMP()
Dim chk As Integer, CHK1 As String
chk = InStr(Trim(str(Val(M_FLX_AMT))), ".")
If chk <> 0 Then
    CHK1 = Mid(Trim(str(Val(M_FLX_AMT))), chk + 1, Len(Trim(Val(M_FLX_AMT))))
    If Len(Trim(CHK1)) = 1 Then
        CMP = Val(CHK1) * 100
    ElseIf Len(Trim(CHK1)) = 2 Then
        CMP = Val(CHK1) * 10
    ElseIf Len(Trim(CHK1)) = 3 Then
        CMP = Val(CHK1) * 1
    End If
Else
    CMP = 0
End If
End Function

