VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_QST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settlement"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "M_FRM_QST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Settlement Details"
      TabPicture(0)   =   "M_FRM_QST.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MS"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "M_INV_IDY"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_QST.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label49"
      Tab(1).Control(1)=   "list4"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton Command1 
         Caption         =   "&Generate"
         Height          =   375
         Left            =   6120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox M_INV_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   4455
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1200
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   16711680
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6800
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Challan Date"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Salesman "
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Van Challan ID"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   675
         Width           =   1065
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Van Challan ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   -74880
         TabIndex        =   3
         Top             =   840
         Width           =   6975
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
            Picture         =   "M_FRM_QST.frx":047A
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_QST.frx":09BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_QST.frx":0D86
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_QST.frx":10DA
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_QST.frx":1712
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_QST.frx":1A66
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_QST.frx":1EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_QST.frx":21BA
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_QST.frx":25C6
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_QST.frx":26D2
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_QST.frx":2A26
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_QST.frx":2E3A
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
      Width           =   7335
      _ExtentX        =   12938
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
            Object.Visible         =   0   'False
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
            Object.ToolTipText     =   "Prefix Area ID"
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
   End
End
Attribute VB_Name = "M_FRM_QST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub INVOICELIST()
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select inv_idy,inv_dat,SMN_NME,SMN_AD4 from SMN,VAN where SMN_nme like '" & Trim(M_INV_IDY) & "%' and VAN.smn_idy=smn.smn_idy " & qry & " GROUP BY INV_IDY ", CON, adOpenKeyset, adLockOptimistic
list4.ListItems.Clear
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET(0)))
        ls.ListSubItems.Add , , Format(REC_SET(1), "dd/MmM/yyyy")
        ls.ListSubItems.Add , , Trim(REC_SET("smn_NME") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("smn_AD4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    Label49 = "Select Van Challan Document"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Van Challan Number Not Found...!", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        M_INV_IDY.Enabled = True
        SendKeys "{HOME}+{END}"
        M_INV_IDY.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click()
Dim RSINL As New ADODB.Recordset
If Len(Trim(M_INV_IDY)) = 0 Then
    MsgBox "Select Van Challan ", vbExclamation
    Exit Sub
End If
RSINL.Open "SELECT SUM(SPD_AMT),SUM(CSD_AMT),SUM(TAX_AMT),SUM(EXD_AMT),SUM(PKG_CST) FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then
    inldiscount = Val(RSINL(0) & "") + Val(RSINL(1) & "")
    inltax = Val(RSINL(2) & "") + Val(RSINL(3) & "") + Val(RSINL(4) & "")
End If
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(DIS_AMT),SUM(ADD_AMT),SUM(TAX_AMT) FROM INV WHERE INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "') AND STK_TYP IN('Salable','Free')", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then
    invdiscount = Val(RSINL(0) & "") + Val(RSINL(1) & "")
    invtax = Val(RSINL(2) & "")
End If
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(FRT_AMT) FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "' and rtrim(ltrim(TXN_RM2))<>'Scheme Amount'", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then inldiscount = Val(inldiscount) + Val(RSINL(0) & "")
MS.TextMatrix(5, 2) = convert(Val(inldiscount) + Val(invdiscount))
MS.TextMatrix(6, 2) = convert(Val(inltax) + Val(invtax))
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT sum((prd_SDR /sub_unt) * tot_pcs) FROM INV,BRD WHERE BRD.BAT_IDY LIKE INV.BAT_IDY AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "') AND INV.STK_TYP IN('Salable','Free')", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then
    MS.TextMatrix(1, 2) = convert(RSINL(0))
End If

'freeqty
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT sum((prd_SDR /sub_unt) * tot_pcs) FROM INV,BRD WHERE INV.BAT_IDY LIKE BRD.BAT_IDY AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "') AND STK_TYP IN('Free') and sch_idy IS NULL", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then MS.TextMatrix(2, 2) = convert(Val(RSINL(0)))
'schqty
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT * FROM lnv,brd,prd WHERE lnv.INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "') and lev_des like brd.bat_idy and brd.prd_idy like prd.prd_idy", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then
    For A = 1 To RSINL.RecordCount
        b = Val(b) + ((Val(RSINL!prd_sdr) / Val(RSINL!sub_unt)) * Val(RSINL!INV_GRS))
        RSINL.MoveNext
    Next
    MS.TextMatrix(3, 2) = convert(Val(b))
    If RSINL.State = 1 Then RSINL.Close
    RSINL.Open "SELECT sum(lnv_amt) FROM lnv WHERE lnv.INV_IDY IN(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "') and (len(alltrim(lev_des))=0 or inv_grs = 0) ", CON, adOpenKeyset, adLockOptimistic
    If RSINL.RecordCount > 0 Then c = Val(RSINL(0) & "")
    MS.TextMatrix(4, 2) = convert(Val(c))
End If
'CASH
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(TXN_AMT) FROM VCH,ACT WHERE VCH.DBT_IDY LIKE ACT.ACT_IDY AND ACT.GRP_CLF LIKE 'Cash Book' and inv_idy in(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "')", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then MS.TextMatrix(8, 2) = convert(Val(RSINL(0) & ""))
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(TXN_AMT) FROM VCH,ACT WHERE VCH.DBT_IDY LIKE ACT.ACT_IDY AND ACT.GRP_CLF LIKE 'Bank Book' and inv_idy in(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "')", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then MS.TextMatrix(9, 2) = convert(Val(RSINL(0) & ""))
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(TXN_AMT) FROM VCH WHERE inv_idy in(SELECT INV_IDY FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "')", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then MS.TextMatrix(10, 2) = convert(Val(RSINL(0) & ""))
If RSINL.State = 1 Then RSINL.Close
RSINL.Open "SELECT SUM(net_amt) FROM INL WHERE ORD_IDY LIKE '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
If RSINL.RecordCount > 0 Then MS.TextMatrix(10, 2) = convert(Val(RSINL(0) & "") - Val(MS.TextMatrix(10, 2)))
CALC
End Sub

Private Sub Form_Activate()
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then Shell "c:\windows\calc.exe"
If KeyCode = 112 Then M_HELP
If KeyCode > 111 And KeyCode < 120 Then If TypeOf M_FRM_QST.ActiveControl Is TextBox Then M_FRM_QST.ActiveControl.Text = FUNKEY(KeyCode)
If KeyCode = 80 And Shift = 2 Then PRINTRECORD
If KeyCode = 88 And Shift = 2 Then Unload Me
If KeyCode = 81 And Shift = 2 Then Cancelrecord
End Sub

Private Sub Form_Load()
MSS.ColAlignment(1) = 7
MSS.ColAlignment(2) = 7
MSS.ColAlignment(3) = 7
MSS.TextMatrix(0, 1) = "Quick Settlement"
MSS.TextMatrix(0, 2) = "Final Settlement"
MSS.TextMatrix(0, 3) = "Diffrence"
MSS.ColWidth(0) = 2000
MSS.ColWidth(1) = 1700
MSS.ColWidth(2) = 1700
MSS.ColWidth(3) = 1500
MSS.Row = 0
MSS.Col = 1
MSS.CellFontBold = True
MSS.Col = 2
MSS.CellFontBold = True
MSS.Col = 3
MSS.CellFontBold = True
MSS.Rows = 1
MSS.Rows = 13
MSS.TextMatrix(1, 0) = "Sales"
MSS.TextMatrix(2, 0) = "Free Quantity        (-)"
MSS.TextMatrix(3, 0) = "Scheme Quantity   (-)"
MSS.TextMatrix(4, 0) = "Scheme Less        (-)"
MSS.TextMatrix(5, 0) = "Discount              (-)"
MSS.TextMatrix(6, 0) = "Taxes & Others    (+)"
MSS.TextMatrix(7, 0) = "Net Sales"
MSS.TextMatrix(8, 0) = "Cash"
MSS.TextMatrix(9, 0) = "Cheques"
MSS.TextMatrix(10, 0) = "Credits"
MSS.TextMatrix(11, 0) = "Variance"
MSS.TextMatrix(12, 0) = "Expenses             (-)"
MSS.Col = 0
For A = 1 To 12
MSS.Row = A
MSS.CellFontBold = True
Next
MSS.Row = 11
MSS.Col = 1
MSS.CellFontBold = True
MSS.Col = 2
MSS.CellFontBold = True
MSS.Col = 3
MSS.CellFontBold = True
For A = 1 To 12
MSS.TextMatrix(A, 1) = "0.00"
MSS.TextMatrix(A, 2) = "0.00"
MSS.TextMatrix(A, 3) = "0.00"
Next
End Sub

Private Sub Form_Keypress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list4.Visible = True Then
        list4.Visible = False
        SSTab1.Tab = 0
        M_INV_IDY.SetFocus
    Else
        Unload Me
    End If
End If
End Sub

Private Sub list4_DblClick()
list4_KeyPress 13
End Sub

Private Sub list4_KeyPress(k As Integer)
If k = 13 Then
M_INV_IDY = Trim(list4.SelectedItem)
DISPDETAILS
End If
End Sub

Private Sub M_INV_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Challan Number"
End Sub
Private Sub M_INV_IDY_KeyDown(k As Integer, s As Integer)
If k = 13 Or k = 40 Then
If Len(Trim(M_INV_IDY)) = 0 Then
    INVOICELIST
Else
    DISPDETAILS
End If
End If
End Sub
Private Sub M_INV_IDY_KeyPress(k As Integer)
If k = 39 Or k = 41 Then k = 0
If k > 96 And k < 123 Then k = k - 32
If M_INV_IDY = "" And k = 27 Then Unload Me
End Sub
Private Sub DISPDETAILS()
Dim rsvan As New ADODB.Recordset
If Len(Trim(M_INV_IDY)) = 10 Then
    rsvan.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.bat_IDY LIKE brd.bat_IDY AND inv_idy like '" & Trim(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
Else
    rsvan.Open "select * from van,PRD,BRD WHERE PRD.PRD_IDY LIKE BRD.PRD_IDY AND VAN.bat_IDY LIKE brd.bat_IDY AND inv_idy like 'RSI" & Tincrement(M_INV_IDY) & "'", CON, adOpenKeyset, adLockOptimistic
End If
If rsvan.RecordCount > 0 Then
M_INV_IDY = Trim(rsvan!INV_IDY)
M_INV_IDY.Enabled = False
For A = 1 To rsvan.RecordCount
    totval = Val(totval) + Val(RATCAL(QTY(Val(rsvan!tot_qty)), Val(rsvan!PRD_rat), Val(rsvan!sub_unt)))
    RetVal = Val(RetVal) + Val(RATCAL(QTY(Val(rsvan!ret_qty)), Val(rsvan!PRD_rat), Val(rsvan!sub_unt)))
    freval = Val(freval) + Val(RATCAL(QTY(Val(rsvan!fre_qty)), Val(rsvan!PRD_rat), Val(rsvan!sub_unt)))
    rsvan.MoveNext
Next
rsvan.MoveFirst
MSS.TextMatrix(1, 1) = convert(Val(totval) - Val(RetVal))
MSS.TextMatrix(2, 1) = convert(Val(freval))
MSS.TextMatrix(5, 1) = convert(Val(rsvan!DIS_AMT))
MSS.TextMatrix(12, 1) = convert(Val(rsvan!EXP_AMT))
MSS.TextMatrix(6, 1) = convert(Val(rsvan!ADD_AMT & ""))
MSS.TextMatrix(8, 1) = convert(Val(rsvan!CAS_AMT))
MSS.TextMatrix(9, 1) = convert(Val(rsvan!CHQ_AMT))
CALC
SSTab1.Tab = 0
list4.Visible = False
MSS.SetFocus
Else
MsgBox "Van Challan ID Not found", vbExclamation
M_INV_IDY.SetFocus
SendKeys "{home}+{end}"
End If
End Sub
Private Sub CALC()
MS.TextMatrix(1, 3) = convert(Val(MS.TextMatrix(1, 1)) - Val(MS.TextMatrix(1, 2)))
For A = 2 To 5
b = Val(b) + Val(MS.TextMatrix(A, 1))
c = Val(c) + Val(MS.TextMatrix(A, 2))
MS.TextMatrix(A, 3) = convert(Val(MS.TextMatrix(A, 1)) - Val(MS.TextMatrix(A, 2)))
Next
MS.TextMatrix(6, 3) = convert(Val(MS.TextMatrix(6, 1)) - Val(MS.TextMatrix(6, 2)))
MS.TextMatrix(7, 1) = convert(Val(MS.TextMatrix(1, 1)) + Val(MS.TextMatrix(6, 1)) - Val(b))
MS.TextMatrix(7, 2) = convert(Val(MS.TextMatrix(1, 2)) + Val(MS.TextMatrix(6, 2)) - Val(c))
MS.TextMatrix(11, 1) = convert(Val(MS.TextMatrix(7, 1)) - (Val(MS.TextMatrix(8, 1)) + Val(MS.TextMatrix(9, 1)) + Val(MS.TextMatrix(10, 1))))
MS.TextMatrix(11, 2) = convert(Val(MS.TextMatrix(7, 2)) - (Val(MS.TextMatrix(8, 2)) + Val(MS.TextMatrix(9, 2)) + Val(MS.TextMatrix(10, 2))))
MS.TextMatrix(7, 3) = convert(Val(MS.TextMatrix(7, 1)) - Val(MS.TextMatrix(7, 2)))
MS.TextMatrix(8, 3) = convert(Val(MS.TextMatrix(8, 1)) - Val(MS.TextMatrix(8, 2)))
MS.TextMatrix(9, 3) = convert(Val(MS.TextMatrix(9, 1)) - Val(MS.TextMatrix(9, 2)))
MS.TextMatrix(10, 3) = convert(Val(MS.TextMatrix(10, 1)) - Val(MS.TextMatrix(10, 2)))
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "c" Then
Cancelrecord
ElseIf Button.KEY = "h" Then
M_HELP
ElseIf Button.KEY = "ed" Then
INVOICELIST
ElseIf Button.KEY = "p" Then
PRINTRECORD
ElseIf Button.KEY = "e" Then
Unload Me
End If
End Sub
Private Sub Cancelrecord()
For A = 1 To MSS.Rows - 1
    MSS.TextMatrix(A, 1) = "0.00"
    MSS.TextMatrix(A, 2) = "0.00"
    MSS.TextMatrix(A, 3) = "0.00"
Next
M_INV_IDY.Enabled = True
M_INV_IDY = ""
M_INV_IDY.SetFocus
End Sub
Private Sub PRINTRECORD()
If Len(Trim(M_INV_IDY)) > 0 Then
    Dim rs As New ADODB.Recordset
    CON.Execute "delete from stl"
    rs.Open "select * from stl", CON, adOpenKeyset, adLockOptimistic
    For A = 1 To MS.Rows - 1
    rs.AddNew
    rs!txn_rmk = Trim(MS.TextMatrix(A, 0))
    rs!rct_Qty = Val(MS.TextMatrix(A, 1))
    rs!ISS_QTY = Val(MS.TextMatrix(A, 2))
    rs!INV_IDY = Trim(M_INV_IDY)
    rs!PRD_IDY = ""
    rs.Update
    Next
    rs.Close
    Openreport CURDIR & "\dos_vch.rp1", "Van Settlement", 0, "", 6
Else
    MsgBox "Select Van Challan", vbExclamation
    M_INV_IDY.SetFocus
End If
End Sub
