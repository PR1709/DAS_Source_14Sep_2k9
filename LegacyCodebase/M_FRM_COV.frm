VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_COV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Conversion"
   ClientHeight    =   5595
   ClientLeft      =   2205
   ClientTop       =   2055
   ClientWidth     =   8790
   Icon            =   "M_FRM_COV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Invoice Details"
      TabPicture(0)   =   "M_FRM_COV.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "MS"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "M_INV_DAT"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CMD_LOAD"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "M_INV_IDY"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "stk_typ"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "&View"
      TabPicture(1)   =   "M_FRM_COV.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "V_LST_CAP"
      Tab(1).Control(1)=   "lv"
      Tab(1).ControlCount=   2
      Begin VB.ListBox stk_typ 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         ItemData        =   "M_FRM_COV.frx":047A
         Left            =   5520
         List            =   "M_FRM_COV.frx":0484
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox M_INV_IDY 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   330
         Left            =   1395
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton CMD_LOAD 
         Caption         =   "&Get SKUs"
         Height          =   375
         Left            =   7365
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker M_INV_DAT 
         Height          =   330
         Left            =   3360
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16580611
         CurrentDate     =   37099
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   4035
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   960
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   7117
         _Version        =   393216
         Rows            =   0
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         GridLines       =   2
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
      Begin MSComctlLib.ListView lv 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Van  ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Salesman Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Place"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   4920
         TabIndex        =   10
         Top             =   555
         Width           =   360
      End
      Begin VB.Label V_LST_CAP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -74760
         TabIndex        =   9
         Top             =   720
         Width           =   8295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2880
         TabIndex        =   7
         Top             =   555
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Number"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   555
         Width           =   1200
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":0499
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":09DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":0DA5
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":10F9
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":1731
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":1A85
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":1EE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":21D9
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":25E5
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":26F1
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":2A45
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":2E59
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_COV.frx":33A5
            Key             =   ""
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   "Conversion Report"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.ToolTipText     =   "System Generated Invoice Number"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "crt"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "brw"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
Attribute VB_Name = "M_FRM_COV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idnew As Boolean, MOD1 As Boolean, r As Integer, V1 As Single, V2 As Single
Private Sub CMD_LOAD_Click()
Dim RSBRD As New ADODB.Recordset
Dim RSINV As New ADODB.Recordset
Dim RSINV1 As New ADODB.Recordset
MS.Cols = 16
MS.Rows = 2
MS.FixedRows = 1
With MS
    .TextMatrix(0, 0) = "Product Name"
    .ColWidth(0) = 2000
    .TextMatrix(0, 1) = "Present Qty"
    .ColAlignment(1) = 7
    .ColWidth(1) = 1000
    .TextMatrix(0, 2) = "Convrt. Qty"
    .ColAlignment(2) = 7
    SCl MS, 2
    .ColWidth(2) = 1000
    .TextMatrix(0, 3) = "Diff Qty"
    .ColAlignment(3) = 7
    .ColWidth(3) = 1000
    .TextMatrix(0, 4) = "Type"
    .ColWidth(4) = 1000
    .TextMatrix(0, 5) = "Product ID"
    .ColWidth(5) = 0
    .TextMatrix(0, 6) = "Batch ID"
    .ColWidth(6) = 0
    .TextMatrix(0, 7) = "Purchase Rate"
    .ColAlignment(7) = 7
    .ColWidth(7) = 1200
    .TextMatrix(0, 8) = "Sale Rate (B)"
    .ColAlignment(8) = 7
    .ColWidth(8) = 1200
    .TextMatrix(0, 9) = "    MRP"
    .ColAlignment(9) = 7
    .ColWidth(9) = 1000
    .TextMatrix(0, 10) = "Batch Name"
    .ColWidth(10) = 2000
    .TextMatrix(0, 11) = "Mfg.Date"
    .ColAlignment(11) = 1
    .ColWidth(11) = 1200
    .TextMatrix(0, 12) = "BBD-Exp.Date"
    .ColAlignment(12) = 1
    .ColWidth(12) = 1200
    .TextMatrix(0, 13) = "Sub Units"
    .ColWidth(13) = 0
    .ColWidth(14) = 0
    .ColWidth(15) = 0
End With
If stk_typ.Text = "Salable" Then
    If RSBRD.State = 1 Then RSBRD.Close
    If idnew = True Then
        RSBRD.Open "SELECT BRD.PRD_IDY,BRD.BAT_IDY,BRD.PRD_PDR,BRD.PRD_SDR,BRD.PRD_MRP,BRD.BAT_NME,BRD.BRD_DOM,BRD.BRD_DOE,BRD.PRD_QTY,PRD.PRD_NME,PRD.STK_TYP,PRD.SUB_UNT FROM BRD,PRD WHERE BRD.PRD_QTY>0 AND PRD.STK_TYP='Salable' and prd.prd_idy=brd.prd_idy ORDER BY PRD_NME", con, adOpenKeyset, adLockPessimistic
    Else
        RSBRD.Open "SELECT BRD.PRD_IDY,BRD.BAT_IDY,BRD.PRD_PDR,BRD.PRD_SDR,BRD.PRD_MRP,BRD.BAT_NME,BRD.BRD_DOM,BRD.BRD_DOE,INV.OPN_QTY,PRD.PRD_NME,PRD.STK_TYP,PRD.SUB_UNT,INV.STK_TYP,INV.PRD_QTY FROM inv,BRD,PRD WHERE BRD.PRD_QTY>0 AND INV.STK_TYP='Damage' and prd.prd_idy=brd.prd_idy and inv.prd_idy like brd.prd_idy and inv.inv_idy like '" & Trim(M_INV_IDY.Text) & "' group by inv.bat_idy", con, adOpenKeyset, adLockPessimistic
    End If
    MS.Rows = RSBRD.RecordCount + 1
    If stk_typ.ListIndex = 0 Then
        STYP = stk_typ.List(1)
    Else
        STYP = stk_typ.List(0)
    End If
    If MOD1 = True Then
        For r = 1 To MS.Rows
            If RSBRD.EOF = False Then
                MS.TextMatrix(r, 0) = Trim(RSBRD(9))
                MS.TextMatrix(r, 2) = QTY(RSBRD(13))
                MS.TextMatrix(r, 3) = QTY(RSBRD(8))
                MS.TextMatrix(r, 4) = STYP
                MS.TextMatrix(r, 1) = QTY(RSBRD(8))
                MS.TextMatrix(r, 5) = RSBRD(0)
                MS.TextMatrix(r, 6) = RSBRD(1)
                MS.TextMatrix(r, 7) = convert(RSBRD(2))
                MS.TextMatrix(r, 8) = convert(RSBRD(3))
                MS.TextMatrix(r, 9) = convert(RSBRD(4))
                MS.TextMatrix(r, 10) = Trim(RSBRD(5))
                MS.TextMatrix(r, 11) = Format(RSBRD(6), "DD/MMM/YYYY")
                MS.TextMatrix(r, 12) = Format(RSBRD(7), "DD/MMM/YYYY")
                MS.TextMatrix(r, 13) = RSBRD(11)
                MS.TextMatrix(r, 14) = Val(RSBRD!SUB_UNT & "")
                DF
                If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
                RSBRD.MoveNext
            End If
        Next
        If MS.Rows > 1 Then
            MS.HighLight = flexHighlightWithFocus
            MS.SetFocus
            MS.Col = 3
            MS.Row = 1
            Exit Sub
        End If
    End If
    If idnew = True Then
        If Not RSBRD.EOF Then
            For r = 1 To MS.Rows
                If RSBRD.EOF = False Then
                    MS.TextMatrix(r, 0) = Trim(RSBRD(9))
                    MS.TextMatrix(r, 2) = "0.000"
                    MS.TextMatrix(r, 3) = QTY(RSBRD(8))
                    MS.TextMatrix(r, 4) = STYP
                    MS.TextMatrix(r, 1) = QTY(RSBRD(8))
                    MS.TextMatrix(r, 5) = RSBRD(0)
                    MS.TextMatrix(r, 6) = RSBRD(1)
                    MS.TextMatrix(r, 7) = convert(RSBRD(2))
                    MS.TextMatrix(r, 8) = convert(RSBRD(3))
                    MS.TextMatrix(r, 9) = convert(RSBRD(4))
                    MS.TextMatrix(r, 10) = Trim(RSBRD(5))
                    MS.TextMatrix(r, 11) = Format(RSBRD(6), "DD/MMM/YYYY")
                    MS.TextMatrix(r, 12) = Format(RSBRD(7), "DD/MMM/YYYY")
                    MS.TextMatrix(r, 13) = RSBRD(11)
                    MS.TextMatrix(r, 14) = Val(RSBRD!SUB_UNT & "")
                    RSBRD.MoveNext
                End If
            Next
            LESSDAMAGE
            MS.HighLight = flexHighlightWithFocus
            MS.SetFocus
            MS.Col = 2
            MS.Row = 1
            Exit Sub
        End If
    End If
End If
If stk_typ.Text = "Damage" Then
    If RSBRD.State = 1 Then RSBRD.Close
    If idnew = True Then
        RSBRD.Open "SELECT BRD.PRD_IDY,BRD.BAT_IDY,BRD.PRD_PDR,BRD.PRD_SDR,BRD.PRD_MRP,BRD.BAT_NME,BRD.BRD_DOM,BRD.BRD_DOE,sum(tot_pcs),PRD.PRD_NME,PRD.STK_TYP,PRD.SUB_UNT,INV.STK_TYP FROM inv,BRD,PRD WHERE INV.STK_TYP='Damage' and BRD.PRD_QTY >0 AND prd.prd_idy=brd.prd_idy and inv.prd_idy like brd.prd_idy and txn_typ in('SAL','PUR','CNV') group by inv.bat_idy", con, adOpenKeyset, adLockPessimistic
    Else
        RSBRD.Open "SELECT BRD.PRD_IDY,BRD.BAT_IDY,BRD.PRD_PDR,BRD.PRD_SDR,BRD.PRD_MRP,BRD.BAT_NME,BRD.BRD_DOM,BRD.BRD_DOE,sum(tot_pcs),PRD.PRD_NME,PRD.STK_TYP,PRD.SUB_UNT,INV.STK_TYP,OPN_QTY,ISS_QTY FROM inv,BRD,PRD WHERE prd.prd_idy=brd.prd_idy and inv.prd_idy like brd.prd_idy and inv.inv_idy like '" & Trim(M_INV_IDY) & "' group by inv.bat_idy", con, adOpenKeyset, adLockPessimistic
    End If
    MS.Rows = RSBRD.RecordCount + 1
    If stk_typ.ListIndex = 0 Then
        STYP = stk_typ.List(1)
    Else
        STYP = stk_typ.List(0)
    End If
    For r = 1 To MS.Rows
        If RSBRD.EOF = False Then
            MS.TextMatrix(r, 0) = Trim(RSBRD(9))
            If idnew = True Then
                MS.TextMatrix(r, 1) = QTY(STM(RSBRD(8), RSBRD(11)))
                MS.TextMatrix(r, 2) = "0.000"
                MS.TextMatrix(r, 3) = QTY(STM(RSBRD(8), RSBRD(11)))
            Else
                MS.TextMatrix(r, 1) = QTY(RSBRD!OPN_QTY & "")
                MS.TextMatrix(r, 2) = QTY(STM(RSBRD(8), RSBRD(11)))
                MS.TextMatrix(r, 3) = QTY(RSBRD!ISS_QTY & "")
            End If
            MS.TextMatrix(r, 4) = STYP
            MS.TextMatrix(r, 5) = RSBRD(0)
            MS.TextMatrix(r, 6) = RSBRD(1)
            MS.TextMatrix(r, 7) = convert(RSBRD(2))
            MS.TextMatrix(r, 8) = convert(RSBRD(3))
            MS.TextMatrix(r, 9) = convert(RSBRD(4))
            MS.TextMatrix(r, 10) = Trim(RSBRD(5))
            MS.TextMatrix(r, 11) = Format(RSBRD(6), "DD/MMM/YYYY")
            MS.TextMatrix(r, 12) = Format(RSBRD(7), "DD/MMM/YYYY")
            MS.TextMatrix(r, 13) = RSBRD(11)
            MS.TextMatrix(r, 14) = Val(RSBRD!SUB_UNT & "")
            DF
            RSBRD.MoveNext
        End If
    Next
    If idnew = True And RSBRD.RecordCount > 0 Then
        LOADSALABLE
        MS.HighLight = flexHighlightWithFocus
        MS.SetFocus
        MS.Col = 2
        MS.Row = 1
    End If
End If
End Sub

Public Sub LOADSALABLE()
Me.MousePointer = vbHourglass
Dim rs As New ADODB.Recordset
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Salable' and txn_typ in('CNV') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 6) = Trim(rs!BAT_IDY) Then
            MS.TextMatrix(b, 15) = Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Damage' and txn_typ in('STO','ADJ') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 6) = Trim(rs!BAT_IDY) Then
            MS.TextMatrix(b, 15) = Val(MS.TextMatrix(b, 15)) + Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next

For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 15)) <> 0 Then
        FLXQTY = MTS(Val(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 14)))
        MS.TextMatrix(A, 1) = QTY(STM((Val(FLXQTY) - Val(MS.TextMatrix(A, 15))), Val(MS.TextMatrix(A, 14))))
    End If
Next
Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
If M_INV_IDY.Enabled = True Then
    M_INV_IDY.SetFocus
End If
stk_typ.ListIndex = 0
'MS.Visible = False
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_COV.ActiveControl Is TextBox Then M_FRM_COV.ActiveControl.Text = Trim(FUNKEY(K))
If lv.Visible = False Then
    If K = 78 And s = 2 Then
        newrec
    ElseIf K = 83 And s = 2 Then
        save
    ElseIf K = 80 And s = 2 Then
        Printrecord
    ElseIf K = 81 And s = 2 Then
        cancel
    ElseIf K = 88 And s = 2 Then
       ' exitfrm
    ElseIf K = 67 And s = 2 Then
        'If Toolbar1.Buttons(12).Enabled = True Then CREATE
    ElseIf K = 68 And s = 2 Then
        'delete
    ElseIf K = 66 And s = 2 Then
       ' If Toolbar1.Buttons(13).Enabled = True Then BROWSE
    ElseIf K = 69 And s = 2 Then
         If M_FRM_COV.ActiveControl.name = "M_INV_IDY" Then V_INV_LST
        'If M_FRM_COV.ActiveControl.name = "M_INV_IDY" Then marlist
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And s = 2 Then
        'Openreport CURDIR & "\lists.RP1", "Markets list", 0, "", 6
    ElseIf K = 82 And s = 2 Then
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If lv.Visible = True Then
        lv.Visible = False
        SSTab1.Tab = 0
        SSTab1.TabEnabled(1) = False
        M_INV_IDY.Enabled = True
        M_INV_IDY.SetFocus
    Else
        If M_FRM_COV.ActiveControl.name = "M_INV_IDY" Then
            Unload Me
        ElseIf M_FRM_COV.ActiveControl.name = "SSTab1" Then
        ElseIf M_FRM_COV.ActiveControl.name = "MS" Then
            MS.HighLight = flexHighlightNever
            SSTab1.Tab = 0
            SSTab1.TabEnabled(1) = False
            save
        Else
            save
        End If
    End If
End If
End Sub

Private Sub Form_Load()
SETYEAR Me
End Sub

Private Sub LV_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If lv.SortOrder = lvwAscending Then
    lv.SortOrder = lvwDescending
ElseIf lv.SortOrder = lvwDescending Then
    lv.SortOrder = lvwAscending
End If
lv.SortKey = CH.Index - 1
End Sub

Private Sub LV_DblClick()
lv_KeyDown 13, 0
End Sub

Private Sub lv_KeyDown(K As Integer, Shift As Integer)
Dim RSINV As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
If K = 13 Then
    If lv.ListItems.count > 0 Then
        RSINV.Open "SELECT * FROM INV WHERE INV_IDY LIKE '" & lv.SelectedItem & "'", con, adOpenKeyset, adLockPessimistic
        If MOD1 = True Then
            M_INV_IDY = lv.SelectedItem
            M_INV_DAT.Value = RSINV("INV_DAT")
            If Trim(RSINV!stk_typ) = "Damage" Then
                stk_typ.ListIndex = 0
            Else
                stk_typ.ListIndex = 1
            End If
            stk_typ.Enabled = False
            lv.Visible = False
            SSTab1.TabEnabled(1) = False
            SSTab1.Tab = 0
            M_INV_IDY.Enabled = False
            CMD_LOAD.SetFocus
        End If
    End If
End If
End Sub
Private Sub M_INV_DAT_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    If stk_typ.Enabled = True Then
        stk_typ.SetFocus
    Else
        CMD_LOAD.SetFocus
    End If
End If
End Sub

Private Sub M_INV_IDY_KeyDown(K As Integer, Shift As Integer)
'If K = 13 Or K = 40 Then
'If Trim(M_INV_IDY) = "" Then
'    v_inv_lst
'Else
''    DISPDETAILS
'End If
'End If
Dim RSINV As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
If K = 13 Or K = 40 Then
If RSINV.State = 1 Then RSINV.Close
    RSINV.Open "select DISTINCT(INV_IDY),INV_DAT,STK_TYP from inv where inv_idy like '" & Sincrement(M_INV_IDY) & "'", con, adOpenStatic
    'RSPRD.Open "SELECT INV.STK_TYP,PRD.PRD_IDY,INV.INV_IDY,INV.PRD_IDY FROM PRD,INV WHERE PRD.PRD_IDY LIKE INV.PRD_IDY AND INV.INV_IDY LIKE '" & Sincrement(M_INV_IDY) & "'", CON, adOpenKeyset, adlockpessimistic
    If RSINV.RecordCount > 0 Then
        M_INV_IDY = Sincrement(M_INV_IDY)
        M_INV_DAT.Value = RSINV("INV_DAT")
        If Trim(RSINV("STK_TYP")) = "Damage" Then
            stk_typ.ListIndex = 0
        Else
            stk_typ.ListIndex = 1
        End If
        stk_typ.Enabled = False
        M_INV_IDY.Enabled = False
        M_INV_DAT.SetFocus
        MOD1 = True
    Else
        V_INV_LST
    End If
End If
End Sub

Private Sub M_INV_IDY_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 47 And K < 58 Or K = 8 Or K = 13 Then
ElseIf K > 96 And K < 123 Then
    K = K - 32
ElseIf K > 64 And K < 91 Then
Else
    K = 0
End If
If M_INV_IDY = "" And K = 27 Then Unload Me
End Sub

Public Sub V_INV_LST()
Dim RSINV As New ADODB.Recordset
Dim ls As ListItem
If RSINV.State = 1 Then RSINV.Close
RSINV.Open "SELECT DISTINCT(INV_IDY),INV_DAT,STK_TYP FROM INv WHERE TXN_TYP LIKE 'CNV' AND INV_IDY LIKE '" & Trim(M_INV_IDY) & "%'", con, adOpenKeyset, adLockPessimistic
If RSINV.RecordCount > 0 Then
    With lv
        .Visible = True
        .ColumnHeaders.CLEAR
        .ColumnHeaders.Add , , "ID", 2000
        .ColumnHeaders.Add , , "Voucher Date", 2000
        .ListItems.CLEAR
        While Not RSINV.EOF
            Set ls = .ListItems.Add(, , RSINV(0))
            ls.ListSubItems.Add , , RSINV(1)
            RSINV.MoveNext
        Wend
    End With
    V_LST_CAP = "Select Voucher"
    SSTab1.TabEnabled(1) = True
    SSTab1.Tab = 1
    lv.SetFocus
    MOD1 = True
Else
    MsgBox "Voucher Not Found !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    M_INV_IDY.Enabled = True
    M_INV_IDY.SetFocus
    MOD1 = False
End If
End Sub
Private Sub MS_KeyPress(K As Integer)
If MS.Row < 0 Then Exit Sub
If MS.TextMatrix(MS.Row, 0) <> "" And MS.Row > 0 Then
    If MS.Col = 2 Then
        If K > 47 And K < 58 Or K = 46 Then
            Text1 = ""
            Text1.MaxLength = 9
            Text1.Visible = True
            Text1.Top = MS.Top + MS.CellTop
            Text1.Left = MS.CellLeft + MS.Left
            Text1.Width = MS.CellWidth
            Text1.Height = MS.CellHeight
            Text1.Visible = True
            Text1.SetFocus
            Text1 = Chr(K)
            Text1.SelStart = Len(Text1)
        ElseIf K = 13 Then
            If MS.TextMatrix(MS.Rows - 1, 0) <> "" Then
                MS.AddItem ""
                MS.Col = 0
                MS.Row = MS.Rows - 1
            End If
        Else
            K = 0
        End If
    End If
End If
End Sub

Private Sub MS_Scroll()
If Text1.Visible = True Then
    Text1 = ""
    Text1.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub stk_typ_KeyPress(K As Integer)
If K = 13 Then
    CMD_LOAD.SetFocus
End If
End Sub
Private Sub newrec()
cancel
M_INV_IDY.Enabled = False
M_INV_DAT.SetFocus
SSTab1.Tab = 0
stk_typ.Enabled = True
idnew = True
MOD1 = False
End Sub
Private Sub cancel()
clrctr M_FRM_COV
M_INV_IDY.Enabled = True
idnew = False
SSTab1.Tab = 0
M_INV_IDY.SetFocus
MOD1 = False
stk_typ.Enabled = True
stk_typ.ListIndex = 0
MS.Rows = 1
End Sub

Private Sub TEXT1_KeyDown(K As Integer, Shift As Integer)
Dim FS As String
If K = 13 Then
    b = InStr(1, Text1, ".")
'    B1 = InStr(1, MS.TextMatrix(MS.Row, 1), ".")
    d = 0
    If b <> 0 Then
        C = Mid(Text1, b, Len(Text1))
        d = Val(C) * 1000
    End If
    'If B1 <> 0 Then
     '   C1 = Mid(Text1, B, Len(Text1))
     '   D1 = Val(C1) * 1000
    'End If
    If MTS(MS.TextMatrix(MS.Row, 1), MS.TextMatrix(MS.Row, 13)) >= MTS(QTY(Text1), MS.TextMatrix(MS.Row, 13)) Then
        If d < Val(MS.TextMatrix(MS.Row, 13)) Then
            Text1.Visible = False
            If Len(Trim(Text1)) = 0 Then
                MS.Text = "0.000"
            Else
                MS.Text = QTY(Text1.Text)
            End If
            MS.SetFocus
         '   MsgBox MS.TextMatrix(MS.Row, 13)
        Else
            MsgBox "Sub Units are " & MS.TextMatrix(MS.Row, 13), vbExclamation, "DAS Version 4.0"
            Text1.Visible = True
            Text1.SetFocus
            SendKeys "{HOME}" + "{END}", 1
            Exit Sub
        End If
    ElseIf Val(Mid(MS.TextMatrix(MS.Row, 1), 1, 3)) < Val(Mid(Trim(Text1), 1, 3)) Then
        MsgBox "No.Of Cases are " & MS.TextMatrix(MS.Row, 1), vbExclamation, "DAS Version 4.0"
        Text1.Visible = True
        Text1.SetFocus
        SendKeys "{HOME}" + "{END}", 1
        Exit Sub
    Else
        MsgBox "Sub Units are " & Mid(Trim(MS.TextMatrix(MS.Row, 1)), 4, 4), vbExclamation, "DAS Version 4.0"
        Text1.Visible = True
        Text1.SetFocus
        SendKeys "{HOME}" + "{END}", 1
        Exit Sub
    End If
    Dim DIFF, DIFF1, DIF
    Dim V1
    Dim V2
    Dim V3
    V1 = MTS(Val(MS.TextMatrix(MS.Row, 1)), Val(MS.TextMatrix(MS.Row, 13)))
    If Not Len(Trim(Text1)) = 0 Then
        V2 = MTS(Text1, MS.TextMatrix(MS.Row, 13))
    End If
    V3 = Val(MS.TextMatrix(MS.Row, 13))
    'V1 = Val(MS.TextMatrix(MS.Row, 1))
    'V2 = Val(Text1)
    DIFF = (V1 - V2)
    DIFF = Int(DIFF / V3)
'    If DIFF <> 0 Then
'        KK = InStr(1, DIFF, ".")
'        KK1 = Mid(STR(DIFF), 1, KK)
'    End If
    DIFF1 = (V1 - V2) Mod Val(MS.TextMatrix(MS.Row, 13))
    DIF = Val(DIFF) + (Val(DIFF1) / 1000)

        '    MsgBox DIFF
        FS = str(Trim(DIF))
        MS.TextMatrix(MS.Row, 3) = QTY(FS)
        'MS.TextMatrix(MS.Row, 3) = STR(Trim(DIF))
    'MS.TextMatrix(MS.Row, 3) = STR(Trim(DIFF))
'    If Mid(Trim(MS.TextMatrix(MS.Row, 3)), 1, 1) = "." Then
 '       MS.TextMatrix(MS.Row, 3) = QTY(Mid(Trim(MS.TextMatrix(MS.Row, 3)), 1, 4))
  '  Else
  If Val(MS.TextMatrix(MS.Row, 3)) = 0 Then
        MS.TextMatrix(MS.Row, 3) = "0.000"
    End If
    MS.ColAlignment(3) = 7
    MS.SetFocus
    If MS.Row < MS.Rows - 1 Then
        MS.Row = MS.Row + 1
    End If
    End If
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 47 And K < 58 Then
ElseIf K = 8 Or K = 46 Then
Else
    K = 0
End If
If K = 39 Or K = 41 Then K = 0
If K = 46 Then
    For T = 1 To Len(Trim(Text1))
        If Mid(Trim(Text1), T, 1) = "." Then
        K = 0
        Exit For
        End If
    Next
    Text1.MaxLength = Len(Trim(Text1)) + 4
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    newrec
ElseIf Button.KEY = "s" Then
    save
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    V_INV_LST
ElseIf Button.KEY = "d" Then
ElseIf Button.KEY = "pre" Then
ElseIf Button.KEY = "brw" Then
ElseIf Button.KEY = "p" Then
    Printrecord
ElseIf Button.KEY = "pr" Then
ElseIf Button.KEY = "crt" Then
End If
End Sub
Private Sub save()
Dim RSINV As New ADODB.Recordset
Dim RSINL As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
If MsgBox("Confirm Record Save ?", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
If MS.Rows = 1 Then
    MsgBox "Load SKU's !", vbExclamation, "DAS Version 4.0"
    CMD_LOAD.SetFocus
    Exit Sub
End If
If Trim(M_INV_IDY) <> "" Then
    MsgBox "Changes cannot be Saved", vbExclamation
    If check("PRT", "PRT_NME", "M_FRM_COV") = True Then
        OpenDocument CURDIR & "\Vouch.RP1", "Stock Conversion Voucher", 2, Trim(M_INV_IDY), Trim(M_INV_IDY), 6, CURDIR & "\inv.cdx,c,inv_idy"
    End If
    cancel
    
    Exit Sub
End If
If RSINV.State = 1 Then RSINV.Close
If RSINL.State = 1 Then RSINL.Close
RSINV.Open "select * from inv", con, adOpenKeyset, adLockPessimistic
RSINL.Open "select * from inl", con, adOpenKeyset, adLockPessimistic
For r = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(r, 2)) > 0 Then
        invid = increment("idy_con")
        Exit For
    End If
Next
IDY = 0
For r = 1 To MS.Rows - 1
If rsprd.State = 1 Then rsprd.Close
rsprd.Open "select * from prd where prd_idy like '" & MS.TextMatrix(r, 5) & "'", con, adOpenKeyset, adLockPessimistic
    If Val(MS.TextMatrix(r, 2)) > 0 Then
    IDY = 1
        RSINV.AddNew
        RSINV("inv_idy") = invid
        For K = 1 To RSINV.Fields.count - 1
            If RSINV.Fields(K).Type = adChar Or RSINV.Fields(K).Type = adVarChar Or RSINV.Fields(K).Type = adBSTR Or RSINV.Fields(K).Type = adLongVarChar Or RSINV.Fields(K).Type = adLongVarWChar Or RSINV.Fields(K).Type = adVarWChar Or RSINV.Fields(K).Type = adBinary Then
                RSINV.Fields(K) = ""
            ElseIf RSINV.Fields(K).Type = adBigInt Or RSINV.Fields(K).Type = adInteger Or RSINV.Fields(K).Type = adNumeric Or RSINV.Fields(K).Type = adSingle Or RSINV.Fields(K).Type = adSmallInt Or RSINV.Fields(K).Type = adTinyInt Or RSINV.Fields(K).Type = adDouble Then
                RSINV.Fields(K) = 0
            ElseIf RSINV.Fields(K).Type = adDate Or RSINV.Fields(K).Type = adDBDate Or RSINV.Fields(K).Type = adDBTimeStamp Then
                RSINV.Fields(K) = Date$
            End If
        Next
        RSINV("INV_DAT") = M_INV_DAT.Value
        RSINV("PRD_IDY") = MS.TextMatrix(r, 5)
        RSINV("BAT_IDY") = MS.TextMatrix(r, 6)
        RSINV("STK_TYP") = MS.TextMatrix(r, 4)
        RSINV("opn_QTY") = Val(MS.TextMatrix(r, 1))
        RSINV("PRD_QTY") = Val(MS.TextMatrix(r, 2))
        RSINV("iss_QTY") = Val(MS.TextMatrix(r, 3))
        RSINV("TXN_TYP") = "CNV"
        RSINV("sub_unt") = MS.TextMatrix(r, 13)
        PU = InStr(1, MS.TextMatrix(r, 2), ".")
        RSINV("prd_unt") = Trim(Mid(Trim(MS.TextMatrix(r, 2)), 1, PU))
        RSINV("prd_pcs") = Mid(Trim(MS.TextMatrix(r, 2)), PU + 1, 3)
        RSINV("tot_pcs") = MTS(MS.TextMatrix(r, 2), MS.TextMatrix(r, 13))
        RSINV("prd_uom") = QTY((rsprd("Con_fac") / rsprd("sub_unt")) * RSINV("tot_pcs"))
        RSINV.Update
    End If
Next
If IDY = 1 Then
    RSINL.AddNew
    RSINL("inv_idy") = invid
    MsgBox "New Voucher ID is " & invid, vbExclamation, "DAS Version 4.0"
    For K = 1 To RSINL.Fields.count - 1
        If RSINL.Fields(K).Type = adChar Or RSINL.Fields(K).Type = adVarChar Or RSINL.Fields(K).Type = adBSTR Or RSINL.Fields(K).Type = adLongVarChar Or RSINL.Fields(K).Type = adLongVarWChar Or RSINL.Fields(K).Type = adVarWChar Or RSINL.Fields(K).Type = adBinary Then
            RSINL.Fields(K) = ""
        ElseIf RSINL.Fields(K).Type = adBigInt Or RSINL.Fields(K).Type = adInteger Or RSINL.Fields(K).Type = adNumeric Or RSINL.Fields(K).Type = adSingle Or RSINL.Fields(K).Type = adSmallInt Or RSINL.Fields(K).Type = adTinyInt Or RSINL.Fields(K).Type = adDouble Then
            RSINL.Fields(K) = 0
        ElseIf RSINL.Fields(K).Type = adDate Or RSINL.Fields(K).Type = adDBDate Or RSINL.Fields(K).Type = adDBTimeStamp Then
            RSINL.Fields(K) = Date$
        End If
    Next
    RSINL("inv_dat") = M_INV_DAT.Value
    RSINL("txn_typ") = "CNV"
    RSINL("SYS_DAT") = Date$
    RSINL("SYS_TIM") = Time$
    RSINL("USR_IDY") = Trim(USERNAME)
    RSINL.Update
End If
If check("PRT", "PRT_NME", "M_FRM_COV") = True Then
    OpenDocument CURDIR & "\Vouch.RP1", "Stock Conversion Voucher", 2, Trim(invid), Trim(invid), 6, CURDIR & "\inv.cdx,c,inv_idy"
End If
cancel
M_INV_IDY.Enabled = True
idnew = False
SSTab1.Tab = 0
M_INV_IDY.SetFocus
MOD1 = False
End Sub
Public Sub DF()
Dim DIFF, DIFF1, DIF
Dim V1, V2, V3, FS As String
If MS.Row = 0 Then MS.Row = 1
    V1 = MTS(Val(MS.TextMatrix(MS.Row, 1)), Val(MS.TextMatrix(MS.Row, 13)))
    If Not Len(Trim(MS.TextMatrix(MS.Row, 2))) = 0 Then V2 = MTS(MS.TextMatrix(MS.Row, 2), MS.TextMatrix(MS.Row, 13))
    V3 = Val(MS.TextMatrix(MS.Row, 13))
    DIFF = (V1 - V2)
    DIFF = Int(DIFF / V3)
    DIFF1 = (V1 - V2) Mod Val(MS.TextMatrix(MS.Row, 13))
    DIF = Val(DIFF) + (Val(DIFF1) / 1000)
    FS = str(Trim(DIF))
    MS.TextMatrix(MS.Row, 3) = QTY(FS)
    If Val(MS.TextMatrix(MS.Row, 3)) = 0 Then MS.TextMatrix(MS.Row, 3) = "0.000"
    MS.ColAlignment(3) = 7
'If MS.Row < MS.Rows - 1 Then MS.Row = MS.Row + 1
End Sub

Public Sub LESSDAMAGE()
Me.MousePointer = vbHourglass
Dim rs As New ADODB.Recordset
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Damage' and txn_typ in('PUR','SAL','CNV') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 6) = Trim(rs!BAT_IDY) Then
            MS.TextMatrix(b, 15) = Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Damage' and txn_typ in('STO','ADJ') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 6) = Trim(rs!BAT_IDY) Then
            MS.TextMatrix(b, 15) = Val(MS.TextMatrix(b, 15)) - Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next
ADDCONVSALABLE
For A = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(A, 15)) > 0 Then
        FLXQTY = MTS(Val(MS.TextMatrix(A, 1)), Val(MS.TextMatrix(A, 14)))
        MS.TextMatrix(A, 1) = QTY(STM((Val(FLXQTY) - Val(MS.TextMatrix(A, 15))), Val(MS.TextMatrix(A, 14))))
    End If
    DF
Next
DF
Me.MousePointer = vbDefault
End Sub

Public Sub ADDCONVSALABLE()
Me.MousePointer = vbHourglass
Dim rs As New ADODB.Recordset
rs.Open "SELECT SUM(TOT_PCS),BAT_IDY,SUB_UNT FROM INV WHERE STK_TYP LIKE 'Salable' and txn_typ in('CNV') group by BAT_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    For b = 1 To MS.Rows - 1
        If MS.TextMatrix(b, 6) = Trim(rs!BAT_IDY) Then
            MS.TextMatrix(b, 15) = Val(MS.TextMatrix(b, 15)) - Val(rs(0) & "")
        End If
    Next
    rs.MoveNext
Next
End Sub

Public Sub Printrecord()
If check("PRT", "PRT_NME", "M_FRM_COV") = True Then
M_FRM_PRV.Check3.Value = 1
Else
M_FRM_PRV.Check3.Value = 0
End If
M_FRM_PRV.Tag = "COV_FRM"
M_FRM_PRV.Show 1
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
M_FRM_DSN.Tag = stk_typ & " CNV"
M_FRM_DSN.Show 1
End Sub
