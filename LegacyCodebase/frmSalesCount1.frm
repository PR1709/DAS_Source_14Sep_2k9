VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSalescount1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales - Count"
   ClientHeight    =   5625
   ClientLeft      =   1455
   ClientTop       =   330
   ClientWidth     =   8820
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSalesCount1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   5760
      TabIndex        =   2
      Top             =   720
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   8705
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmSalesCount1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label21"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "M_ACT_NME"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "M_SUP_NME"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "M_INV_DAT"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "M_INV_IDY"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "M_CRT_IDY"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "M_DBT_IDY"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmSalesCount1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSF"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmSalesCount1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox M_DBT_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2070
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox M_CRT_IDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2070
         MaxLength       =   10
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2955
         Width           =   1215
      End
      Begin VB.TextBox M_INV_IDY 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   330
         Left            =   2070
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker M_INV_DAT 
         Height          =   330
         Left            =   5580
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   75038723
         CurrentDate     =   37099
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   4395
         Left            =   -74880
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   7752
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
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
      Begin VB.Label M_SUP_NME 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   3360
         TabIndex        =   12
         Top             =   2520
         Width           =   3405
      End
      Begin VB.Label M_ACT_NME 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   3390
         TabIndex        =   11
         Top             =   2955
         Width           =   3405
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Account"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   840
         TabIndex        =   10
         Top             =   3030
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   990
         TabIndex        =   9
         Top             =   1515
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   4950
         TabIndex        =   8
         Top             =   1515
         Width           =   345
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Id"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1350
         TabIndex        =   7
         Top             =   2595
         Width           =   540
      End
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
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
            Picture         =   "frmSalesCount1.frx":0496
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":09DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":0DA2
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":10F6
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":172E
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":1A82
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":1EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":21D6
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":25E2
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":26EE
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":2A42
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":2E56
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSalesCount1.frx":33A2
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
      Width           =   8820
      _ExtentX        =   15558
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
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pre"
            Object.ToolTipText     =   "Click Here for System Generated Invoice Number"
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
Attribute VB_Name = "frmSalescount1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msfcol As Integer
Dim msfrow As Integer



Private Sub EXT_REC()
Set Form = Nothing
Unload Me
End Sub


Private Sub M_INV_DAT_CloseUp()
M_INV_DAT.SetFocus
End Sub
Private Sub M_INV_DAT_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Select Date Using Arrow Keys Or Press F4"
End Sub

Private Sub M_INV_DAT_KeyDown(KeyCode As Integer, s As Integer)
If KeyCode = 13 Then
    If M_VAN_IDY.Visible = True Then
        If M_VAN_IDY.Enabled = True Then
            M_VAN_IDY.SetFocus
        Else
            M_CST_IDY.SetFocus
        End If
    Else
        M_CST_IDY.SetFocus
    End If
End If
End Sub



Private Sub Form_Activate()
checkdongle
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If
If KeyCode > 111 And KeyCode < 120 Then If TypeOf frmSalesCount.ActiveControl Is TextBox Then frmSalesCount.ActiveControl.Text = FUNKEY(KeyCode)
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then

End If
End Sub
Private Sub Form_Load()
checkdongle
MOD_REC = False
FillGrid
Text5.Locked = False
Text5.Enabled = True
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub M_INV_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press Ctrl+N to Generate New Invoice ID by System, Press <Enter> to View Existing Invoice"
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
Set Form = Nothing
End Sub






Private Sub M_INV_IDY_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
If Len(Trim(M_INV_IDY)) = 0 Then
    V_INV_LST
Else
    inV = M_INV_IDY
    m_smn = Text2
    clrctr M_FRM_SAL
    Text2 = m_smn
    M_CST_NME = ""
    M_CST_AD1 = ""
    M_CST_AD2 = ""
    M_CST_AD3 = ""
    M_CST_AD4 = ""
    Label18 = "0"
    Label19 = "0.00"
    Label20 = "0.00"
    M_GRS_AMT = "0.00"
    M_NET_AMT = "0.00"
    M_FFR_NME = ""
    M_ACT_NME = ""
    M_INV_IDY = inV
    DISPDETAILS
End If
End If
End Sub

Private Sub MSF_EnterCell()
'If MSF.Col = 0 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> for Product ID View"
'ElseIf MSF.Col = 1 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Stock Type"
'ElseIf MSF.Col = 2 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Enter Product Quantity"
'ElseIf MSF.Col = 4 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Rate"
'ElseIf MSF.Col = 3 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Unit of Measure (Nos, Kgs, Ltrs, Mtrs, MT)"
'ElseIf MSF.Col = 5 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Amount"
'ElseIf MSF.Col = 6 Then
'    M_MDI.STS_BAR.Panels(1).Text = "MRP"
'ElseIf MSF.Col = 7 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Product Name"
'ElseIf MSF.Col = 8 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Short Name"
'ElseIf MSF.Col = 9 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Sub Units"
'ElseIf MSF.Col = 10 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Unit of Measurement"
'ElseIf MSF.Col = 15 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Batch"
'ElseIf MSF.Col = 16 Then
'    M_MDI.STS_BAR.Panels(1).Text = "Tax %"
'End If

End Sub

Private Sub MSF_GotFocus()
    If MSF.Col = 3 Or MSF.Col = 4 Then
        Text5.Visible = True
        Text5.Width = MSF.CellWidth
        Text5.Top = MSF.Top + MSF.CellTop
        Text5.Height = MSF.RowHeight(MSF.Row)
        Text5.Left = MSF.Left + MSF.CellLeft
        Text5.Text = MSF.TextMatrix(MSF.Row, MSF.Col)
        Text5.SelStart = 1
        Text5.Visible = True
        Text5.SetFocus
    End If
End Sub

Private Sub MSF_keypress(K As Integer)
    If K = 13 Then
        CalculateValues
        If MSF.Col = 3 Then
            MSF.Col = 4
            Text5.Visible = True
            Text5.Width = MSF.CellWidth
            Text5.Top = MSF.Top + MSF.CellTop
            Text5.Height = MSF.RowHeight(MSF.Row)
            Text5.Left = MSF.Left + MSF.CellLeft
            Text5.Text = MSF.TextMatrix(MSF.Row, MSF.Col)
            Text5.SelStart = 1
            Text5.Visible = True
            Text5.SetFocus
        ElseIf MSF.Col = 4 Then
            MSF.Col = 3
            If MSF.Row > 0 And MSF.Row < MSF.Rows - 1 Then
                MSF.Row = MSF.Row + 1
            End If
            Text5.Visible = True
            Text5.Width = MSF.CellWidth
            Text5.Top = MSF.Top + MSF.CellTop
            Text5.Height = MSF.RowHeight(MSF.Row)
            Text5.Left = MSF.Left + MSF.CellLeft
            Text5.Text = MSF.TextMatrix(MSF.Row, MSF.Col)
            Text5.SelStart = 1
            Text5.Visible = True
            Text5.SetFocus
        End If
    ElseIf K = 38 Then
            CalculateValues
            If MSF.Row > 1 Then
                MSF.Row = MSF.Row - 1
            End If
            Text5.Visible = True
            Text5.Width = MSF.CellWidth
            Text5.Top = MSF.Top + MSF.CellTop
            Text5.Height = MSF.RowHeight(MSF.Row)
            Text5.Left = MSF.Left + MSF.CellLeft
            Text5.Text = MSF.TextMatrix(MSF.Row, MSF.Col)
            Text5.SelStart = 1
            Text5.Visible = True
            Text5.SetFocus
    ElseIf K = 40 Then
            CalculateValues
            If MSF.Row < MSF.Rows - 1 Then
                MSF.Row = MSF.Row + 1
            End If
            Text5.Visible = True
            Text5.Width = MSF.CellWidth
            Text5.Top = MSF.Top + MSF.CellTop
            Text5.Height = MSF.RowHeight(MSF.Row)
            Text5.Left = MSF.Left + MSF.CellLeft
            Text5.Text = MSF.TextMatrix(MSF.Row, MSF.Col)
            Text5.SelStart = 1
            Text5.Visible = True
            Text5.SetFocus
    
    End If
End Sub

Private Sub MSF_LeaveCell()
If Text5.Visible = True Then Text5.Visible = False
End Sub

Private Sub MSF_Scroll()
If Text5.Visible = True Then
    Text5 = ""
    Text5.Visible = False
    MSF.SetFocus
End If
End Sub










Private Sub TEXT5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MSF.Col = 3 Or MSF.Col = 4 Then
        MSF.TextMatrix(MSF.Row, MSF.Col) = Text5.Text
    End If
    MSF_keypress 13
ElseIf KeyCode = 40 Then
    MSF.TextMatrix(MSF.Row, MSF.Col) = Text5.Text
    MSF_keypress 40
ElseIf KeyCode = 38 Then
    MSF.TextMatrix(MSF.Row, MSF.Col) = Text5.Text
    MSF_keypress 38
End If
End Sub

Private Sub Text5_KeyPress(K As Integer)
'If K = 39 Or K = 41 Or K = 37 Then K = 0
'K = CHKNUM(K)
End Sub

'Private Sub Text5_LostFocus()
'CalculateValues
'Text5.Visible = False
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
'    newrecord
ElseIf Button.KEY = "s" Then
    'saverecord
ElseIf Button.KEY = "e" Then
EXT_REC
ElseIf Button.KEY = "p" Then
'Printrecord
ElseIf Button.KEY = "c" Then
'CANCELRECORD
ElseIf Button.KEY = "pre" Then
ElseIf Button.KEY = "AUS" Then
    If Button.Value = tbrPressed Then
        M_AUS_OK = 1
    Else
        M_AUS_OK = 0
    End If

ElseIf Button.KEY = "h" Then
M_HELP
ElseIf Button.KEY = "pre" Then
If CHECKTB("prt.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        
        Exit Sub
End If
Picture1.Visible = True
If check("prt", "prt_nme", "M_FRM_SAL") = False Then
Check3.Value = 0
Else
Check3.Value = 1
End If
Check3.SetFocus
Toolbar1.Enabled = False
SSTab1.Enabled = False
End If
End Sub







Private Sub InitializeGrid()
    With MSF
        .TextMatrix(0, 0) = "Prd. Name"
        .TextMatrix(0, 1) = "Batch Name"
        .TextMatrix(0, 2) = "Stock"
        .ColAlignment(2) = 7
        .TextMatrix(0, 3) = "+"
        SCl MSF, 3
        .ColAlignment(3) = 7
        .TextMatrix(0, 4) = "-"
        SCl MSF, 4
        .ColAlignment(4) = 7
        .TextMatrix(0, 5) = "Balance"
        .ColAlignment(5) = 7
        .TextMatrix(0, 6) = "Product Id"
        .TextMatrix(0, 7) = "Batch Id"
    End With
    MSF.Rows = 2
End Sub

Private Sub FillGrid()
InitializeGrid
Dim rec As New ADODB.Recordset
rec.Open "select prd.prd_idy , prd.prd_nme, prd.cur_lev, brd.bat_nme, bat_idy from prd, brd where trim(prd.prd_idy) = trim(brd.prd_idy)", con, adOpenKeyset, adLockPessimistic

While Not rec.EOF
    With MSF
        .TextMatrix(MSF.Rows - 1, 0) = Trim(rec!PRD_NME)
        .TextMatrix(MSF.Rows - 1, 1) = Trim(rec!Bat_nme)
        .TextMatrix(MSF.Rows - 1, 2) = Trim(rec!CUR_LEV)
        .TextMatrix(MSF.Rows - 1, 6) = Trim(rec!prd_idy)
        .TextMatrix(MSF.Rows - 1, 7) = Trim(rec!BAT_IDY)
    End With
    rec.MoveNext
    If Not rec.EOF Then MSF.Rows = MSF.Rows + 1
Wend
End Sub

Public Sub CalculateValues()
For i = 0 To MSF.Rows - 1
    MSF.TextMatrix(i, 5) = Val(MSF.TextMatrix(i, 2)) + Val(MSF.TextMatrix(i, 3)) - Val(MSF.TextMatrix(i, 4))
Next
End Sub
