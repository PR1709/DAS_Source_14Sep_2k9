VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form M_FRM_SIC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commissions"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   Icon            =   "M_FRM_SIC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab2 
      Height          =   5160
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   9102
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Commissions"
      TabPicture(0)   =   "M_FRM_SIC.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SMNNME"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "mscl"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "MSC"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "MSI"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SMNID"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txttemp"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lstRates"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ILSTRATES"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ITXTtEMP"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TGTID"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCol"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "&View"
      TabPicture(1)   =   "M_FRM_SIC.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "lv1"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtCol 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TGTID 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   705
         Width           =   975
      End
      Begin VB.TextBox ITXTtEMP 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2220
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ListBox ILSTRATES 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "M_FRM_SIC.frx":047A
         Left            =   4080
         List            =   "M_FRM_SIC.frx":0487
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2220
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ListBox lstRates 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "M_FRM_SIC.frx":0494
         Left            =   2040
         List            =   "M_FRM_SIC.frx":04A1
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2100
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txttemp 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox SMNID 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1140
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MSI 
         Height          =   3375
         Left            =   1665
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   22
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
      Begin MSFlexGridLib.MSFlexGrid MSC 
         Height          =   3255
         Left            =   1762
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1680
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   22
         FixedCols       =   0
         RowHeightMin    =   330
         BackColorSel    =   14737632
         ForeColorSel    =   8388608
         GridColor       =   8421504
         GridColorFixed  =   4210752
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   2
         ScrollBars      =   2
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
      Begin MSFlexGridLib.MSFlexGrid mscl 
         Height          =   3495
         Left            =   60
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1635
         Visible         =   0   'False
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   22
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
      Begin MSComctlLib.ListView lv1 
         Height          =   3615
         Left            =   -74400
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1095
         Visible         =   0   'False
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6376
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   0
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Select Salesman"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   -74400
         TabIndex        =   16
         Top             =   780
         Width           =   7695
      End
      Begin VB.Label SMNNME 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4320
         TabIndex        =   13
         Top             =   1140
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Commission ID"
         Height          =   195
         Left            =   1950
         TabIndex        =   11
         Top             =   773
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select Salesman"
         Height          =   195
         Left            =   1800
         TabIndex        =   1
         Top             =   1215
         Width           =   1185
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   480
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
            Picture         =   "M_FRM_SIC.frx":04AE
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIC.frx":09F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIC.frx":0DBA
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIC.frx":110E
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIC.frx":1746
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIC.frx":1A9A
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIC.frx":1EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIC.frx":21EE
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIC.frx":25FA
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIC.frx":2706
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIC.frx":2A5A
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SIC.frx":2E6E
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
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
                  Key             =   "a"
                  Text            =   "&a Commission Report"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview"
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
            Object.ToolTipText     =   "Prefix Product ID "
            ImageIndex      =   10
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A"
                  Text            =   "Purchase Rate Check"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "B"
                  Text            =   "Prefix Product ID"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "d"
            Object.ToolTipText     =   "Delete (Ctrl+D)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_FRM_SIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSSMN As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
Dim RSBRD As New ADODB.Recordset
Dim rssic As New ADODB.Recordset, qv As String
Dim rschk As New ADODB.Recordset
Dim r As Integer, C As Integer, am As Single
Dim MD As Boolean
'Private Sub FDATE_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    'TDATE.SetFocus
'End If
'End Sub

Private Sub Form_Activate()
SSTab2.Tab = 0
SSTab2.TabEnabled(1) = False
If TGTID.Enabled = True Then TGTID.SetFocus
MD = False
End Sub

Private Sub Form_KeyDown(K As Integer, S As Integer)
If K = 78 And S = 2 Then
    newrec
ElseIf K = 83 And S = 2 Then
    save
ElseIf K = 81 And S = 2 Then
    cancel
ElseIf K = 80 And S = 2 Then
    'PRINT
ElseIf K = 112 Then
    M_HELP
ElseIf K = 68 And S = 2 Then
    DELREC
ElseIf K = 88 And S = 2 Then
    Unload Me
ElseIf K = 69 And S = 2 Then
    DISPLIST
ElseIf K = 67 And S = 2 Then
        If Toolbar1.Buttons(12).Enabled = True Then CRT_DPR
ElseIf K = 27 Then
    If Me.ActiveControl.name = "LV1" Then
        smnid.SetFocus
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If Me.ActiveControl.name = "TGTID" Then
        Unload Me
    ElseIf SSTab2.Tab = 1 Then
        If Label5.Caption = "Select Salesman" Then
            SSTab2.Tab = 0
            lv1.Visible = False
            smnid.SetFocus
            SSTab2.TabEnabled(1) = False
        Else
            SSTab2.Tab = 0
            lv1.Visible = False
            TGTID.SetFocus
            SSTab2.TabEnabled(1) = False
        End If
    Else
        save
    End If
End If
End Sub

Private Sub Form_Load()
SETYEAR Me
loadflds
End Sub

Private Sub ILSTRATES_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MSI.Visible = True Then
        MSI.Text = Trim(ILSTRATES.Text)
        If MSI.Col <> MSI.Cols - 1 Then
            MSI.Col = MSI.Col + 1
        End If
        ILSTRATES.Visible = False
        MSI.SetFocus
    End If
End If
End Sub

Private Sub inc_typ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Trim(inc_typ.Text) = "Quantity" Then
        qv = "Qty"
    ElseIf Trim(inc_typ.Text) = "Value" Then
        qv = "Value"
    End If
    PRDCAT.Enabled = True
    PRDCAT.SetFocus
    PRDCAT.ListIndex = 0
End If
End Sub

Private Sub INCTYP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Trim(INCTYP.Text) = "Commission" Then
        MSI.Visible = False
        mscl.Visible = False
        MSC.Visible = True
        'inc_typ.Enabled = False
        loadmsc
        MSC.SetFocus
        MSC.Col = 2
    ElseIf Trim(INCTYP.Text) = "Collection" Then
        MSC.Visible = False
        MSI.Visible = False
        mscl.Visible = True
        PRATA.Enabled = True
        loadmscl
'        inc_typ.Enabled = True
'        MSC.Visible = False
'        MSI.Visible = True
'        loadmsi
'        inc_typ.SetFocus
'        inc_typ.ListIndex = 0
    End If
End If
End Sub

Private Sub itxttemp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MSI.Col = 3 Then
        If Val(CMP(ITXTtEMP.Text)) >= Val(MSI.TextMatrix(MSI.Row, 8)) Then
            MsgBox "Sub Units are " & Trim(MSI.TextMatrix(MSI.Row, 8)), vbExclamation, "DAS Version 4.0"
            ITXTtEMP.SetFocus
            SendKeys "{home}+{end}"
        Else
            MSI.Text = QTY(STR(Val(ITXTtEMP.Text)))
            ITXTtEMP.Visible = False
            MSI.Col = MSI.Col + 1
            MSI.SetFocus
        End If
    ElseIf MSI.Col = 4 Then
        If Val(CMP(ITXTtEMP.Text)) >= Val(MSI.TextMatrix(MSI.Row, 8)) Then
            MsgBox "Sub Units are " & Trim(MSI.TextMatrix(MSI.Row, 8)), vbExclamation, "DAS Version 4.0"
            ITXTtEMP.SetFocus
            SendKeys "{home}+{end}"
        Else
            If Val(ITXTtEMP.Text) < Val(MSI.TextMatrix(MSI.Row, 3)) Then
                MsgBox "Maximum Quantity should be greater than Minimum Qty", vbExclamation, "DAS Version 4.0"
                ITXTtEMP.SetFocus
                SendKeys "{home}+{end}"
            Else
                MSI.Text = QTY(STR(Val(ITXTtEMP.Text)))
                ITXTtEMP.Visible = False
                MSI.Col = MSI.Col + 1
                MSI.SetFocus
            End If
        End If
    ElseIf MSI.Col = 5 Then
        If Val(MSI.TextMatrix(MSI.Row, 5)) >= 100 Then
            MsgBox "Invalid Percentage", vbExclamation, "DAS Version 4.0'"
            ITXTtEMP.SetFocus
            SendKeys "{home}+{end}"
        Else
            MSI.Text = AMT(ITXTtEMP.Text)
            ITXTtEMP.Visible = False
            MSI.Col = MSI.Col + 1
            MSI.SetFocus
        End If
    ElseIf MSI.Col = 6 Then
        MSI.Text = AMT(ITXTtEMP.Text)
        ITXTtEMP.Visible = False
        If MSI.Row <> MSI.Rows - 1 Then
            MSI.Row = MSI.Row + 1
            MSI.Col = 2
        End If
        MSI.SetFocus
    End If
End If
End Sub

Private Sub lstRates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MSI.Visible = True Then
        MSI.Text = Trim(lstRates.Text)
        lstRates.Visible = False
        If MSI.Col < 5 Then
            MSI.Col = MSI.Col + 1
        Else
            MSI.Row = MSI.Row + 1
            MSI.Col = 1
        End If
        MSI.SetFocus
    ElseIf MSC.Visible = True Then
        MSC.Text = Trim(lstRates.Text)
        If Trim(lstRates.Text) = "A" Then
            FLD = "prd_pdr"
        ElseIf Trim(lstRates.Text) = "B" Then
            FLD = "prd_sdr"
        ElseIf Trim(lstRates.Text) = "C" Then
            FLD = "prd_mrp"
        Else
            FLD = "prd_sdr"
        End If
        If RSBRD.State = 1 Then RSBRD.Close
        If Trim(lstRates.Text) = "A" Then
            RSBRD.Open "select bas_rat from prd,brd where prd.prd_idy like brd.prd_idy and brd.bat_idy='" & Trim(MSC.TextMatrix(MSC.Row, 8)) & "'", CON, adOpenKeyset, adLockPessimistic
        Else
            RSBRD.Open "select " & FLD & " from brd where bat_idy='" & Trim(MSC.TextMatrix(MSC.Row, 8)) & "'", CON, adOpenKeyset, adLockPessimistic
        End If
        If Not RSBRD.EOF Then
            am = RSBRD(0)
            If Val(MSC.TextMatrix(MSC.Row, 5)) > 0 Then
            MSC.TextMatrix(MSC.Row, 6) = AMT(am * (Val(MSC.TextMatrix(MSC.Row, 5)) / 100))
            End If
        End If
        lstRates.Visible = False
        If MSC.Col < 5 Then
            MSC.Col = MSC.Col + 1
        Else
            MSC.Row = MSC.Row + 1
            MSC.Col = 2
        End If
        MSC.SetFocus
    End If
    MSC.Col = 6
End If
End Sub

Private Sub lv1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If lv1.SortOrder = lvwAscending Then
    lv1.SortOrder = lvwDescending
ElseIf lv1.SortOrder = lvwDescending Then
    lv1.SortOrder = lvwAscending
End If
lv1.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub LV1_DblClick()
LV1_KeyDown 13, 0
End Sub

Private Sub LV1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If KeyCode = 13 Then
    If lv1.ListItems.COUNT > 0 Then
        Dim lst As ListSubItem
        If Trim(Label5.Caption) = "Select Salesman" Then
            If RSSMN.State = 1 Then RSSMN.Close
            Set lst = lv1.SelectedItem.ListSubItems.Item(1)
            RSSMN.Open "select * from smn where smn_idy like '" & Trim(lst.Text) & "'", CON, adOpenKeyset, adLockPessimistic
            smnid = Trim(RSSMN("smn_idy"))
            SMNNME = RSSMN("SMN_NME")
            lv1.Visible = False
            SSTab2.Tab = 0
            SSTab2.TabEnabled(1) = False
            MSI.Visible = False
            mscl.Visible = False
            MSC.Visible = True
            loadmsc
            MSC.SetFocus
            MSC.Col = 2
'            If 'FDATE.Enabled = True Then
'                'FDATE.SetFocus
'            Else
'                If MSI.Visible = True Then
'                    MSI.SetFocus
'                    MSI.Col = 2
'                Else
'                    msc.SetFocus
'                    msc.Col = 2
'                End If
'            End If
        Else
            If rssic.State = 1 Then rssic.Close
            Set lst = lv1.SelectedItem.ListSubItems.Item(1)
            rssic.Open "select sic.* from sic where txn_idy like '" & Trim(lst.Text) & "'", CON, adOpenKeyset, adLockPessimistic
            If rssic.RecordCount > 0 Then
                TGTID.Enabled = False
                TGTID = Trim(rssic("txn_idy"))
                smnid = Trim(rssic("act_idy"))
            End If
            If RSSMN.State = 1 Then RSSMN.Close
            RSSMN.Open "SELECT * FROM SMN WHERE SMN_IDY='" & Trim(smnid) & "'", CON, adOpenKeyset, adLockPessimistic
            If Not RSSMN.EOF Then SMNNME = RSSMN("SMN_NME")
                lmsc
                MSC.Visible = True
                MSI.Visible = False
                mscl.Visible = False
                MSC.Rows = rssic.RecordCount + 1
                r = 1
                If RSBRD.State = 1 Then RSBRD.Close
                RSBRD.Open "select bat_nme,bat_idy from brd", CON, adOpenKeyset, adLockPessimistic
                If rsprd.State = 1 Then rsprd.Close
                rsprd.Open "select prd_idy,prd_nme,sub_unt from prd", CON, adOpenKeyset, adLockPessimistic
                While Not rssic.EOF
                    If rsprd.RecordCount > 0 Then rsprd.MoveFirst
                    rsprd.Find "prd_idy='" & rssic("prd_idy") & "'"
                    If rsprd.EOF = False Then
                        MSC.TextMatrix(r, 0) = rsprd("prd_nme") 'rssic("prd_nme")
                        MSC.TextMatrix(r, 9) = rsprd("sub_unt") 'rssic("sub_unt")
                    End If
                    If RSBRD.RecordCount > 0 Then RSBRD.MoveFirst
                    RSBRD.Find "bat_idy='" & rssic("bat_idy") & "'"
                    If RSBRD.EOF = False Then MSC.TextMatrix(r, 1) = RSBRD("bat_nme")
                    MSC.TextMatrix(r, 2) = rssic("rat_idy")
                    MSC.TextMatrix(r, 3) = QTY(rssic("min_qty"))
                    MSC.TextMatrix(r, 4) = QTY(rssic("min_vol"))
                    MSC.TextMatrix(r, 5) = AMT(rssic("min_pcg"))
                    MSC.TextMatrix(r, 6) = AMT(rssic("min_val"))
                    MSC.TextMatrix(r, 7) = rssic("prd_idy")
                    MSC.TextMatrix(r, 8) = rssic("bat_idy")
                    rssic.MoveNext
                    r = r + 1
                Wend
                If rsprd.State = 1 Then rsprd.Close
                'rsprd.Open "select prd_nme,prd_idy,sub_unt from prd where prd_idy not in(select prd_idy from sic where txn_idy='" & Trim(TGTID) & "')", CON, adOpenKeyset, adlockpessimistic
                rsprd.Open "SELECT PRD_NME,PRD_IDY,SUB_UNT FROM PRD WHERE PRD_IDY NOT IN(SELECT PRD_IDY FROM SIC WHERE TXN_IDY='" & Trim(TGTID) & "')", CON, adOpenKeyset, adLockPessimistic
                MSC.Rows = MSC.Rows + rsprd.RecordCount
                While Not rsprd.EOF
                    MSC.TextMatrix(r, 0) = Trim(rsprd("prd_nme")) & ""
                    MSC.TextMatrix(r, 7) = Trim(rsprd("prd_idy")) & ""
                    MSC.TextMatrix(r, 8) = "" 'Trim(rsprd(2)) & ""  ' BAT IDY
                    MSC.TextMatrix(r, 9) = Trim(rsprd("sub_unt"))
                    MSC.TextMatrix(r, 1) = "" 'Trim(rsprd(3))   ' bATCH NAME
                    MSC.TextMatrix(r, 2) = "B"
                    rsprd.MoveNext
                    r = r + 1
                Wend
                MSC.FixedCols = 2
            SSTab2.Tab = 0
            SSTab2.TabEnabled(1) = False
            lv1.Visible = False
            smnid.SetFocus
            SendKeys "{home}+{end}"
        End If
    End If
End If
End Sub

Private Sub msc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then
    If MSC.Rows > 1 Then
    If MSC.Col = 3 Then
    ElseIf MSC.Col = 4 Then
    ElseIf MSC.Col = 2 Then
        lstRates.Visible = True
        lstRates.Left = MSC.Left + MSC.CellLeft
        lstRates.Top = MSC.Top + MSC.CellTop
        lstRates.Height = MSC.CellHeight
        lstRates.Width = MSC.CellWidth
        lstRates.ListIndex = 0
        lstRates.SetFocus
    ElseIf MSC.Col = 5 Then
    ElseIf MSC.Col = 6 Then
        If Val(Trim(MSC.TextMatrix(MSC.Row, 5))) = 0 Then
            txtTemp.Text = ""
            txtTemp.Visible = True
            txtTemp.Left = MSC.Left + MSC.CellLeft
            txtTemp.Top = MSC.Top + MSC.CellTop
            txtTemp.Height = MSC.CellHeight
            'txttemp.Width = msc.CellWidth
            txtTemp.Width = 1250
            txtTemp = Chr(KeyAscii)
            txtTemp.SetFocus
            txtTemp.SelStart = Len(Trim(txtTemp))
        End If
    End If
    End If
End If
End Sub

Private Sub mscl_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Then
    If mscl.Col = 3 Then
        If mscl.Row > 1 Then
            If Val(mscl.TextMatrix(mscl.Row, 0)) > 0 Then
                If Val(mscl.TextMatrix(mscl.Row, 2)) = 0 Then
                    txtCol.Visible = True
                    txtCol.MaxLength = 12
                    txtCol.Left = mscl.Left + mscl.CellLeft
                    txtCol.Top = mscl.Top + mscl.CellTop
                    txtCol.Height = mscl.CellHeight
                    txtCol.Width = mscl.CellWidth
                    txtCol.Text = Chr(KeyAscii)
                    txtCol.SelStart = Len(Trim(txtCol))
                    txtCol.SetFocus
                End If
            End If
        Else
            If Val(mscl.TextMatrix(mscl.Row, 2)) = 0 Then
                txtCol.Visible = True
                txtCol.MaxLength = 12
                txtCol.Left = mscl.Left + mscl.CellLeft
                txtCol.Top = mscl.Top + mscl.CellTop
                txtCol.Height = mscl.CellHeight
                txtCol.Width = mscl.CellWidth
                txtCol.Text = Chr(KeyAscii)
                txtCol.SelStart = Len(Trim(txtCol))
                txtCol.SetFocus
            End If
        End If
    ElseIf mscl.Col = 2 Then
        If mscl.Row > 1 Then
            If Val(mscl.TextMatrix(mscl.Row, 0)) > 0 Then
                txtCol.Visible = True
                txtCol.MaxLength = 5
                txtCol.Left = mscl.Left + mscl.CellLeft
                txtCol.Top = mscl.Top + mscl.CellTop
                txtCol.Height = mscl.CellHeight
                txtCol.Width = mscl.CellWidth
                txtCol.Text = Chr(KeyAscii)
                txtCol.SelStart = Len(Trim(txtCol))
                txtCol.SetFocus
            End If
        Else
            txtCol.Visible = True
            txtCol.MaxLength = 5
            txtCol.Left = mscl.Left + mscl.CellLeft
            txtCol.Top = mscl.Top + mscl.CellTop
            txtCol.Height = mscl.CellHeight
            txtCol.Width = mscl.CellWidth
            txtCol.Text = Chr(KeyAscii)
            txtCol.SelStart = Len(Trim(txtCol))
            txtCol.SetFocus
        End If
    ElseIf mscl.Col = 1 And mscl.Row > 1 Then
        If Val(mscl.TextMatrix(mscl.Row, 0)) > 0 Then
            txtCol.Visible = True
            txtCol.MaxLength = 12
            txtCol.Left = mscl.Left + mscl.CellLeft
            txtCol.Top = mscl.Top + mscl.CellTop
            txtCol.Height = mscl.CellHeight
            txtCol.Width = mscl.CellWidth
            txtCol.Text = Chr(KeyAscii)
            txtCol.SelStart = Len(Trim(txtCol))
            txtCol.SetFocus
        End If
    Else
        txtCol.Visible = True
        txtCol.MaxLength = 12
        txtCol.Left = mscl.Left + mscl.CellLeft
        txtCol.Top = mscl.Top + mscl.CellTop
        txtCol.Height = mscl.CellHeight
        txtCol.Width = mscl.CellWidth
        txtCol.Text = Chr(KeyAscii)
        txtCol.SelStart = Len(Trim(txtCol))
        txtCol.SetFocus
    End If
ElseIf KeyAscii = 13 Then
    If mscl.Col = 3 Then
        If mscl.Row <> mscl.Rows - 1 Then
            mscl.Row = mscl.Row + 1
            mscl.Col = 0
        End If
    End If
End If
End Sub

Private Sub MSI_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then
    If MSI.Col = 3 Or MSI.Col = 4 Or MSI.Col = 5 Then
        ITXTtEMP.Visible = True
        ITXTtEMP.Text = ""
        ITXTtEMP.Left = MSI.Left + MSI.CellLeft
        ITXTtEMP.Top = MSI.Top + MSI.CellTop
        ITXTtEMP.Height = MSI.CellHeight
        ITXTtEMP.Width = MSI.CellWidth
        ITXTtEMP = Chr(KeyAscii)
        ITXTtEMP.SetFocus
        ITXTtEMP.SelStart = Len(Trim(ITXTtEMP))
    ElseIf MSI.Col = 2 Then
        ILSTRATES.Visible = True
        ILSTRATES.Left = MSI.Left + MSI.CellLeft
        ILSTRATES.Top = MSI.Top + MSI.CellTop
        ILSTRATES.Height = MSI.CellHeight
        ILSTRATES.Width = MSI.CellWidth
        ILSTRATES.ListIndex = 0
        ILSTRATES.SetFocus
    ElseIf MSI.Col = 6 Then
        If Val(MSI.TextMatrix(MSI.Row, 5)) = 0 Then
            ITXTtEMP.Visible = True
            ITXTtEMP.Text = ""
            ITXTtEMP.Left = MSI.Left + MSI.CellLeft
            ITXTtEMP.Top = MSI.Top + MSI.CellTop
            ITXTtEMP.Height = MSI.CellHeight
            ITXTtEMP.Width = MSI.CellWidth
            ITXTtEMP = Chr(KeyAscii)
            ITXTtEMP.SetFocus
            ITXTtEMP.SelStart = Len(Trim(ITXTtEMP))
        End If
    End If
End If
End Sub

Private Sub PRDCAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    saltyp.Enabled = True
    saltyp.SetFocus
    saltyp.ListIndex = 0
End If
End Sub

Private Sub saltyp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MSI.SetFocus
    MSI.Col = 2
End If
End Sub

Private Sub smnid_KeyDown(K As Integer, S As Integer)
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If K = 13 Or K = 40 Then
        If RSSMN.State = 1 Then RSSMN.Close
        RSSMN.Open "select * from smn where smn_idy like '" & Sincrement(smnid) & "'", CON, adOpenKeyset, adLockPessimistic
        If RSSMN.RecordCount > 0 Then
            smnid = Sincrement(smnid)
            SMNNME = RSSMN("SMN_NME")
            MSI.Visible = False
            mscl.Visible = False
            MSC.Visible = True
            loadmsc
            If MSC.Rows > 1 Then MSC.Row = 1
            MSC.SetFocus
            MSC.Col = 2
        Else
            smnlist
        End If
End If
End Sub

Private Sub SMNID_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
If KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TGTID_KeyDown(K As Integer, S As Integer)
If CHECKTB("sic.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If K = 13 Or K = 40 Then
    If rssic.State = 1 Then rssic.Close
'    rssic.Open "select * from sic where txn_idy like '" & Sincrement(TGTID) & "'", CON, adOpenStatic
    If Len(Trim(TGTID)) > 0 Then
        rssic.Open "select sic.* from sic where SIC.txn_idy like '" & Sincrement(Trim(TGTID)) & "'", CON, adOpenKeyset, adLockPessimistic
        If rssic.RecordCount > 0 Then
            TGTID = Sincrement(TGTID)
            TGTID.Enabled = False
            txlist
            smnid.SetFocus
            MD = True
        Else
            MsgBox "Transaction ID Not Found!", vbExclamation, "DAS Version 4.0"
            TGTID.SetFocus
            SendKeys "{home}+{end}"
        End If
    Else
        txnlist
    End If
End If
End Sub

Private Sub TGTID_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
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
    If Me.ActiveControl.name = "SMNID" Then
        DISPLIST
    ElseIf Me.ActiveControl.name = "TGTID" Then
        txnlist
    End If
ElseIf Button.KEY = "d" Then
    DELREC
ElseIf Button.KEY = "crt" Then
    CRT_DPR
End If
End Sub
Private Sub CRT_DPR()
MD = False
save
cancel
End Sub
Private Sub DELREC()
On Error GoTo lab
If CHECKTB("SIC.DBF") = 1 Then
    MsgBox UCase(TNAME) & "  Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If MsgBox("Confirm Commission Rate Deletion ?", vbYesNo + vbCritical, "DAS Version 4.0") = vbYes Then
   CON.Execute "DELETE FROM SIC WHERE TXN_IDY LIKE '" & Trim(TGTID) & "'"
   cancel
Else
    cancel
End If
Exit Sub
lab:
MsgBox err.Description
End Sub

Private Sub DISPLIST()
Dim ls As ListItem
If RSSMN.State = 1 Then RSSMN.Close
RSSMN.Open "SELECT * FROM SMN", CON, adOpenKeyset, adLockPessimistic
If RSSMN.RecordCount > 0 Then
    SSTab2.TabEnabled(1) = True
    SSTab2.Tab = 1
    With lv1
        .ColumnHeaders.CLEAR
        .ColumnHeaders.Add , , "Salesman Name", 3000
        .ColumnHeaders.Add , , "ID", 1500
        .ListItems.CLEAR
        While Not RSSMN.EOF
            Set ls = .ListItems.Add(, , RSSMN("smn_nme"))
            ls.ListSubItems.Add , , RSSMN("smn_idy")
            RSSMN.MoveNext
        Wend
        .Visible = True
        .SetFocus
    End With
End If
End Sub
    
Private Sub loadmsc()
With MSC
    lmsc
    If rsprd.State = 1 Then rsprd.Close
    'rsprd.Open "select prd_nme,prd.prd_idy,brd.bat_idy,brd.bat_nme,sub_unt from prd,brd where prd.prd_idy like brd.prd_idy", CON, adOpenKeyset, adlockpessimistic
    rsprd.Open "select prd_nme,prd_idy,sub_unt from prd where stk_typ in('Salable','Free')", CON, adOpenKeyset, adLockPessimistic
    .Rows = rsprd.RecordCount + 1
    r = 1
    While Not rsprd.EOF
        .TextMatrix(r, 0) = Trim(rsprd("prd_nme")) & ""
        .TextMatrix(r, 7) = Trim(rsprd("prd_idy")) & ""
        .TextMatrix(r, 8) = "" 'Trim(rsprd(2)) & ""
        .TextMatrix(r, 9) = Trim(rsprd("sub_unt"))
        .TextMatrix(r, 1) = "" 'Trim(rsprd(3))
        .TextMatrix(r, 2) = "B"
        rsprd.MoveNext
        r = r + 1
    Wend
    .FixedCols = 2
End With
End Sub
Private Sub loadmsi()
With MSI
    lmsi
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "select prd_nme,prd.prd_idy,brd.bat_idy,brd.bat_nme,sub_unt from prd,brd where prd.prd_idy like brd.prd_idy", CON, adOpenKeyset, adLockPessimistic
    .Rows = rsprd.RecordCount + 1
    r = 1
    While Not rsprd.EOF
        .TextMatrix(r, 0) = Trim(rsprd("prd_nme")) & ""
        .TextMatrix(r, 9) = Trim(rsprd("prd_idy")) & ""
        .TextMatrix(r, 7) = Trim(rsprd(2)) & ""
        .TextMatrix(r, 8) = Trim(rsprd("sub_unt"))
        .TextMatrix(r, 1) = Trim(rsprd(3))
        .TextMatrix(r, 2) = "B"
        .TextMatrix(r, 6) = "0.00"
        rsprd.MoveNext
        r = r + 1
    Wend
    .FixedCols = 2
End With
End Sub
Private Sub newrec()
cancel
smnid.SetFocus
End Sub
Private Sub save()
If MsgBox("Confirm Record save ?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
If RSSMN.State = 1 Then RSSMN.Close
RSSMN.Open "SELECT SMN_NME FROM SMN WHERE SMN_IDY='" & Trim(smnid) & "'", CON, adOpenKeyset
If RSSMN.EOF = True Then
    MsgBox "Salesman Not Found!", vbExclamation, "DAS Version 4.0"
    smnid.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
If rssic.State = 1 Then rssic.Close
rssic.Open "SELECT * FROM SIC WHERE ACT_IDY = '" & Trim(smnid) & "'", CON, adOpenKeyset, adLockPessimistic
    If rssic.RecordCount > 0 And MD = False Then
        MsgBox ("Commission Slab for this Salesman Already Available! Cannot Create Duplicate Slab! "), vbCritical, "DAS Version 4.0"
        cancel
        Exit Sub
    End If
    'If Trim(INCTYP.Text) = "Commission" Then
    If MD = False Then
        For r = 1 To MSC.Rows - 1
            If Val(MSC.TextMatrix(r, 6)) > 0 Then
                Exit For
            End If
        Next
        If r <> MSC.Rows Then
        TGTID = increment("idy_txn")
            For r = 1 To MSC.Rows - 1
'                If Val(MSC.TextMatrix(r, 6)) > 0 Then
                    rssic.AddNew
                    rssic("txn_idy") = Trim(TGTID)
                    rssic("act_idy") = Trim(smnid)
                    'rssic("com_inc") = Trim(Mid(Trim(INCTYP.Text), 1, 1))
                    rssic("frm_dat") = Date$ 'FDATE.Value
                    rssic("end_dat") = Date$ 'TDATE.Value
                    rssic("min_qty") = Val(MSC.TextMatrix(r, 3))
                    rssic("min_pcg") = Val(MSC.TextMatrix(r, 5))
                    rssic("min_val") = Val(MSC.TextMatrix(r, 6))
                    rssic("min_vol") = Val(MSC.TextMatrix(r, 4))
                    rssic("max_qty") = Val(MSC.TextMatrix(r, 3))
                    rssic("max_pcg") = Val(MSC.TextMatrix(r, 5))
                    rssic("max_val") = Val(MSC.TextMatrix(r, 6))
                    rssic("max_vol") = Val(MSC.TextMatrix(r, 4))
                    rssic("prd_idy") = Trim(MSC.TextMatrix(r, 7))
                    rssic("bat_idy") = Trim(MSC.TextMatrix(r, 8))
                    rssic("rat_idy") = Trim(MSC.TextMatrix(r, 2))
'                    rssic("pro_rat") = Val(STR(PRATA.Value))
                    rssic.Update
'                End If
            Next
            MsgBox "New Transaction ID is " & Trim(TGTID), vbExclamation, "DAS Version 4.0"
            cancel
        Else
            MsgBox "Define Per SKU CASE - Commission Amount!", vbExclamation, "DAS Version 4.0"
            MSC.SetFocus
            MSC.Row = 1
            MSC.Col = 3
            Exit Sub
        End If
    Else
        If rssic.State = 1 Then rssic.Close
        'rssic.Open "select * from sic where txn_idy='" & Trim(TGTID) & "'", CON, adOpenKeyset, adlockpessimistic
        rssic.Open "SELECT * FROM SIC", CON, adOpenKeyset, adLockPessimistic
        CON.Execute "DELETE FROM SIC where txn_idy='" & Trim(TGTID) & "'"
        For r = 1 To MSC.Rows - 1
'            If Val(MSC.TextMatrix(r, 6)) > 0 Then
                rssic.AddNew
                rssic("txn_idy") = Trim(TGTID)
                rssic("act_idy") = Trim(smnid)
'                rssic("com_inc") = Trim(Mid(Trim(INCTYP.Text), 1, 1))
                rssic("frm_dat") = Date$ 'FDATE.Value
                rssic("end_dat") = Date$ 'TDATE.Value
                rssic("min_qty") = Val(MSC.TextMatrix(r, 3))
                rssic("min_pcg") = Val(MSC.TextMatrix(r, 5))
                rssic("min_val") = Val(MSC.TextMatrix(r, 6))
                rssic("min_vol") = Val(MSC.TextMatrix(r, 4))
                rssic("max_qty") = Val(MSC.TextMatrix(r, 3))
                rssic("max_pcg") = Val(MSC.TextMatrix(r, 5))
                rssic("max_val") = Val(MSC.TextMatrix(r, 6))
                rssic("max_vol") = Val(MSC.TextMatrix(r, 4))
                rssic("prd_idy") = Trim(MSC.TextMatrix(r, 7))
                rssic("bat_idy") = Trim(MSC.TextMatrix(r, 8))
                rssic("rat_idy") = Trim(MSC.TextMatrix(r, 2))
'                rssic("pro_rat") = Val(STR(PRATA.Value))
                rssic.Update
'            End If
        Next
        cancel
    End If
End Sub

Private Sub cancel()
'FDATE.Enabled = True
'FDATE = Date
'TDATE.Enabled = True
'TDATE = Date
MSI.Rows = 1
MSC.Rows = 1
mscl.Rows = 1
txtCol.Visible = False
smnid.Enabled = True
'PRDCAT.Enabled = True
smnid = ""
SMNNME = ""
'inc_typ.Enabled = True
MD = False
ITXTtEMP.Visible = False
ILSTRATES.Visible = False
txtTemp.Visible = False
lstRates.Visible = False
TGTID.Enabled = True
SSTab2.Tab = 0
SSTab2.TabEnabled(1) = False
TGTID = ""
TGTID.SetFocus
End Sub
Private Sub smnlist()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If RSSMN.State = 1 Then RSSMN.Close
RSSMN.Open "SELECT * FROM SMN where smn_nme like '" & Trim(smnid) & "%'", CON, adOpenKeyset, adLockPessimistic
If RSSMN.RecordCount > 0 Then
    SSTab2.TabEnabled(1) = True
    SSTab2.Tab = 1
    Label5 = "Select Salesman"
    With lv1
        .ColumnHeaders.CLEAR
        .ColumnHeaders.Add , , "Salesman Name", 3000
        .ColumnHeaders.Add , , "ID", 1500
        .ListItems.CLEAR
        While Not RSSMN.EOF
            Set ls = .ListItems.Add(, , RSSMN("smn_nme"))
            ls.ListSubItems.Add , , RSSMN("smn_idy")
            RSSMN.MoveNext
        Wend
        .Visible = True
        .SetFocus
    End With
Else
    If SSTab2.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
        SSTab2.Tab = 0
        SSTab2.TabEnabled(1) = False
        smnid.Enabled = True
        smnid.SetFocus
    End If
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
CON.Execute "delete from scl"
    M_REP_SMN.Tag = "COMREP"
    M_REP_SMN.Show 1
    cancel
End Sub

Private Sub txtCol_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(txtCol)) = 0 Then
        txtCol.Visible = False
        mscl.SetFocus
        Exit Sub
    End If
    If mscl.Row > 1 And mscl.Col = 0 Then
        If Val(txtCol) <= Val(mscl.TextMatrix(mscl.Row - 1, 1)) Then
            MsgBox "Invalid Range!", vbExclamation, "DAS Version 4.0"
            txtCol.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    End If
    If mscl.Col = 1 Then
        If Val(txtCol) < Val(mscl.TextMatrix(mscl.Row, 0)) Then
            MsgBox "Invalid Range!", vbExclamation, "DAS Version 4.0"
            txtCol.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    End If
    If mscl.Col = 2 Then
        If Val(txtCol) >= 100 Then
            MsgBox "Invalid Percentage", vbExclamation, "DAS Version 4.0"
            txtCol.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        Else
            mscl.Text = convert(Trim(Val(txtCol)))
            mscl.TextMatrix(mscl.Row, 3) = "0.00"
        End If
    Else
        mscl.Text = convert(Trim(Val(txtCol)))
    End If
    txtCol = ""
    txtCol.Visible = False
    If mscl.Col = 3 Then
        If mscl.Row <> mscl.Rows - 1 Then
            mscl.Row = mscl.Row + 1
            mscl.Col = 0
            mscl.SetFocus
        Else
            mscl.SetFocus
        End If
    Else
        mscl.Col = mscl.Col + 1
        mscl.SetFocus
    End If
End If
End Sub

Private Sub txttemp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MSC.Visible = True Then
        If MSC.Col = 3 Then
        ElseIf MSC.Col = 4 Then
        Else
            MSC.Text = AMT(Trim(txtTemp.Text))
            txtTemp.Visible = False
            If MSC.Col < 5 Then
                MSC.Col = MSC.Col + 1
            Else
                If MSC.Row <> MSC.Rows - 1 Then MSC.Row = MSC.Row + 1
                MSC.Col = 2
            End If
        End If
        MSC.SetFocus
    End If
End If
End Sub
Private Sub loadflds()
With MSC
    .Cols = 8
    .TextMatrix(0, 0) = "SKU Name"
    .ColAlignment(0) = 1
    .ColWidth(0) = 3000
    .TextMatrix(0, 1) = "Rate - ABC"
    .ColWidth(1) = 1000
    SCl MSC, 1
    .TextMatrix(0, 2) = "Qty Cases"
    .ColWidth(2) = 0
    SCl MSC, 2
    .TextMatrix(0, 3) = "Volume"
    .ColWidth(3) = 0
    SCl MSC, 3
    .TextMatrix(0, 4) = "Percent"
    .ColWidth(4) = 0
    SCl MSC, 4
    .TextMatrix(0, 5) = "Cmn per CASE"
    .ColWidth(5) = 1200
    SCl MSC, 5
    .TextMatrix(0, 6) = "Product ID"
    .ColWidth(6) = 0
    .TextMatrix(0, 7) = "Batch ID"
    .ColWidth(7) = 0
End With
End Sub

Public Function CMP(TT As String)
Dim chk As Integer, CHK1 As String
chk = InStr(Trim(STR(Val(TT))), ".")
If chk <> 0 Then
    CHK1 = Mid(Trim(STR(Val(TT))), chk + 1, Len(Trim(Val(TT))))
    If Len(Trim(CHK1)) = 1 Then
        CMP = Val(CHK1) * 100
    ElseIf Len(Trim(CHK1)) = 2 Then
        CMP = Val(CHK1) * 10
    ElseIf Len(Trim(CHK1)) = 3 Then
        CMP = Val(CHK1) * 1
    End If
End If
End Function

Private Sub txnlist()
Dim ls As ListItem
If CHECKTB("sic.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If rssic.State = 1 Then rssic.Close
'rssic.Open "SELECT sic.txn_idy,sic.frm_dat,sic.end_dat,sic.com_inc,smn.smn_nme FROM Sic,smn where sic.act_idy=smn.smn_idy GROUP BY SIC.txn_IDY", CON, adOpenKeyset, adlockpessimistic
rssic.Open "select * from sic group by txn_idy", CON, adOpenKeyset, adLockPessimistic
If rssic.RecordCount > 0 Then
    SSTab2.TabEnabled(1) = True
    SSTab2.Tab = 1
    Label5 = "Select Commission ID"
    With lv1
        .ColumnHeaders.CLEAR
        .ColumnHeaders.Add , , "Salesman Name", 3000
        .ColumnHeaders.Add , , "Cmsn ID", 1300
        .ListItems.CLEAR
        If RSSMN.State = 1 Then RSSMN.Close
        RSSMN.Open "Select * from smn", CON, adOpenKeyset, adLockPessimistic
        While Not rssic.EOF
            If RSSMN.RecordCount > 0 Then RSSMN.MoveFirst
            RSSMN.Find "smn_idy='" & rssic("act_idy") & "'"
            If RSSMN.EOF = False Then Set ls = .ListItems.Add(, , RSSMN("smn_nme"))             ' Salesman Name
            ls.ListSubItems.Add , , rssic("txn_idy") ' Transaction ID
            rssic.MoveNext
        Wend
        .Visible = True
        .SetFocus
    End With
    MD = True
Else
    If SSTab2.Tab = 0 Then
        MsgBox "Transaction Not Found !", vbExclamation, "DAS Version 4.0"
        SSTab2.Tab = 0
        SSTab2.TabEnabled(1) = False
        TGTID.Enabled = True
        TGTID.SetFocus
        MD = False
    End If
End If
End Sub

Private Sub txlist()
'TGTID = Trim(rssic("txn_idy"))
smnid = Trim(rssic("act_idy"))
If RSSMN.State = 1 Then RSSMN.Close
RSSMN.Open "SELECT SMN_NME FROM SMN WHERE SMN_IDY='" & Trim(smnid) & "'", CON, adOpenKeyset, adLockPessimistic
If Not RSSMN.EOF Then SMNNME = RSSMN("SMN_NME")
    lmsc
    MSC.Visible = True
    MSI.Visible = False
    mscl.Visible = False
    MSC.Rows = rssic.RecordCount + 1
    r = 1
    If RSBRD.State = 1 Then RSBRD.Close
    RSBRD.Open "select bat_nme,bat_idy from brd", CON, adOpenKeyset, adLockPessimistic
    If rsprd.State = 1 Then rsprd.Close
    rsprd.Open "select prd_idy,prd_nme,sub_unt from prd", CON, adOpenKeyset, adLockPessimistic
    While Not rssic.EOF
        If rsprd.RecordCount > 0 Then rsprd.MoveFirst
        rsprd.Find "prd_idy='" & rssic("prd_idy") & "'"
        If rsprd.EOF = False Then
            MSC.TextMatrix(r, 0) = rsprd("prd_nme") 'rssic("prd_nme")
            MSC.TextMatrix(r, 8) = rsprd("sub_unt") 'rssic("sub_unt")
        End If
        If RSBRD.RecordCount > 0 Then RSBRD.MoveFirst
        RSBRD.Find "bat_idy='" & rssic("bat_idy") & "'"
        If RSBRD.EOF = False Then MSC.TextMatrix(r, 1) = RSBRD("bat_nme")
        MSC.TextMatrix(r, 2) = rssic("rat_idy")
        MSC.TextMatrix(r, 3) = QTY(rssic("min_qty"))
        MSC.TextMatrix(r, 4) = QTY(rssic("min_vol"))
        MSC.TextMatrix(r, 5) = AMT(rssic("min_pcg"))
        MSC.TextMatrix(r, 6) = AMT(rssic("min_val"))
        MSC.TextMatrix(r, 7) = rssic("prd_idy")
        MSC.TextMatrix(r, 8) = rssic("bat_idy")
        rssic.MoveNext
        r = r + 1
    Wend
    If rsprd.State = 1 Then rsprd.Close
    'rsprd.Open "select prd_nme,prd_idy,sub_unt from prd where prd_idy not in(select prd_idy from sic where txn_idy='" & Trim(TGTID) & "')", CON, adOpenKeyset, adlockpessimistic
    rsprd.Open "SELECT PRD_NME,PRD_IDY,SUB_UNT FROM PRD WHERE PRD_IDY NOT IN(SELECT PRD_IDY FROM SIC WHERE TXN_IDY='" & Trim(TGTID) & "')", CON, adOpenKeyset, adLockPessimistic
    MSC.Rows = MSC.Rows + rsprd.RecordCount
    While Not rsprd.EOF
        MSC.TextMatrix(r, 0) = Trim(rsprd("prd_nme")) & ""
        MSC.TextMatrix(r, 7) = Trim(rsprd("prd_idy")) & ""
        MSC.TextMatrix(r, 8) = "" 'Trim(rsprd(2)) & ""  ' BAT IDY
        MSC.TextMatrix(r, 9) = Trim(rsprd("sub_unt"))
        MSC.TextMatrix(r, 1) = "" 'Trim(rsprd(3))   ' bATCH NAME
        MSC.TextMatrix(r, 2) = "B"
        rsprd.MoveNext
        r = r + 1
    Wend
    MSC.FixedCols = 2
    MSC.Visible = True
    MSI.Visible = False
    smnid.SetFocus
    SendKeys "{home}+{end}"
End Sub
Private Sub lmsc()
With MSC
    .Visible = True
    MSI.Visible = False
    .Cols = 10
    .TextMatrix(0, 0) = "SKU Name"
    .ColWidth(0) = 2500
    .TextMatrix(0, 1) = "Batch Name"
    '.ColWidth(1) = 1500
    .ColWidth(1) = 0
    .ColAlignment(1) = 1
    .TextMatrix(0, 2) = "Rate - ABC"
    SCl MSC, 2
    .ColWidth(2) = 1260
    .TextMatrix(0, 3) = "Qty Cases"
    .ColWidth(3) = 0
    .ColAlignment(3) = 7
    SCl MSC, 3
    .TextMatrix(0, 4) = "Volume"
    .ColWidth(4) = 0
    .ColAlignment(4) = 7
    SCl MSC, 4
    .TextMatrix(0, 5) = "Percent"
    .ColWidth(5) = 0
    .ColAlignment(5) = 7
    SCl MSC, 5
    .TextMatrix(0, 6) = "Cmn Per CASE"
    .ColWidth(6) = 1200
    .ColAlignment(6) = 7
    SCl MSC, 6
    .TextMatrix(0, 7) = "Product ID"
    .ColWidth(7) = 0
    .TextMatrix(0, 8) = "Batch ID"
    .ColWidth(8) = 0
    .TextMatrix(0, 9) = "Sub Units"
    .ColWidth(9) = 0
End With
End Sub
Private Sub lmsi()
With MSI
    .Visible = True
    MSC.Visible = False
    .Cols = 10
    .TextMatrix(0, 0) = "SKU Name"
    .ColWidth(0) = 2500
    .ColAlignment(0) = 1
    .TextMatrix(0, 1) = "Batch Name"
    .ColWidth(1) = 1500
    .ColAlignment(1) = 1
    .TextMatrix(0, 2) = "Rate - ABC"
    .ColWidth(2) = 800
    .ColAlignment(2) = 1
    '.TextMatrix(0, 2) = "Qty .PC"
    '.ColWidth(2) = 1200
    '.ColAlignment(2) = 7
    .TextMatrix(0, 3) = "Min." & qv
    .ColWidth(3) = 900
    .ColAlignment(3) = 7
    .TextMatrix(0, 4) = "Max." & qv
    .ColWidth(4) = 900
    .ColAlignment(4) = 7
    .TextMatrix(0, 5) = "Percent"
    .ColWidth(5) = 800
    .ColAlignment(5) = 7
    .TextMatrix(0, 6) = "Cmn Per CASE"
    .ColWidth(6) = 1200
    .ColAlignment(6) = 7
    .TextMatrix(0, 7) = "Batch ID"
    .ColWidth(7) = 0
    .TextMatrix(0, 8) = "Sub Units"
    .ColWidth(8) = 0
    .TextMatrix(0, 9) = "Product ID"
    .ColWidth(9) = 0
End With
End Sub
Private Sub loadmscl()
With mscl
    .Cols = 4
    .Rows = 6
    .TextMatrix(0, 0) = "Min.Range"
    .ColWidth(0) = 2000
    .ColAlignment(0) = 7
    SCl mscl, 0
    .TextMatrix(0, 1) = "Max.Range"
    .ColWidth(1) = 2000
    .ColAlignment(1) = 7
    SCl mscl, 1
    .TextMatrix(0, 2) = "Percent"
    .ColWidth(2) = 2000
    .ColAlignment(2) = 7
    SCl mscl, 2
    .TextMatrix(0, 3) = "Cmn Per CASE"
    .ColWidth(3) = 2000
    .ColAlignment(3) = 7
    SCl mscl, 3
    .Visible = True
    .SetFocus
    .Col = 1
    .Row = 1
    mscl.TextMatrix(1, 0) = "0.00"
End With
End Sub
