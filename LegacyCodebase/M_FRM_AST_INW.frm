VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_AST_INW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asset Inward Entry"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "M_FRM_AST_INW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5140
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Asset Details"
      TabPicture(0)   =   "M_FRM_AST_INW.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ASTNME"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "PTYNME"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TXNIDY"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TXNDAT"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ASTIDY"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "astqty"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ms"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ptyidy"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "text3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_AST_INW.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "list4"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.TextBox text3 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6000
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox ptyidy 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid ms 
         Height          =   4575
         Left            =   5265
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   8070
         _Version        =   393216
         RowHeightMin    =   300
         BackColorSel    =   14737632
         ForeColorSel    =   16711680
         FocusRect       =   0
         GridLines       =   2
         ScrollBars      =   2
      End
      Begin VB.TextBox astqty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox ASTIDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker TXNDAT 
         Height          =   330
         Left            =   1200
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74317827
         CurrentDate     =   37736
      End
      Begin VB.TextBox TXNIDY 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   480
         Width           =   1215
      End
      Begin MSComctlLib.ListView list4 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1065
         Visible         =   0   'False
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6376
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label20 
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
         Left            =   -74520
         TabIndex        =   16
         Top             =   720
         Width           =   7815
      End
      Begin VB.Label PTYNME 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2520
         TabIndex        =   14
         Top             =   1440
         Width           =   2565
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Party ID"
         Height          =   195
         Left            =   555
         TabIndex        =   12
         Top             =   1485
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
         Height          =   195
         Left            =   540
         TabIndex        =   9
         Top             =   2475
         Width           =   585
      End
      Begin VB.Label ASTNME 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2520
         TabIndex        =   8
         Top             =   1920
         Width           =   2565
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Asset ID"
         Height          =   195
         Left            =   525
         TabIndex        =   6
         Top             =   1965
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Left            =   780
         TabIndex        =   5
         Top             =   1035
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inward ID"
         Height          =   195
         Left            =   435
         TabIndex        =   3
         Top             =   555
         Width           =   690
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   600
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
            Picture         =   "M_FRM_AST_INW.frx":047A
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_AST_INW.frx":09BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_AST_INW.frx":0D86
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_AST_INW.frx":10DA
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_AST_INW.frx":1712
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_AST_INW.frx":1A66
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_AST_INW.frx":1EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_AST_INW.frx":21BA
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_AST_INW.frx":25C6
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_AST_INW.frx":26D2
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_AST_INW.frx":2A26
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_AST_INW.frx":2E3A
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
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
                  Text            =   "Assets Inward Entry"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "act"
                  Text            =   "&Accounts List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "dlr"
                  Text            =   "&Dealers List"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sup"
                  Text            =   "&Supplier List"
               EndProperty
            EndProperty
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
            Object.ToolTipText     =   "Prefix Account ID "
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
Attribute VB_Name = "M_FRM_AST_INW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD1 As Boolean, indidy
Private Sub ASTIDY_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub ASTIDY_KeyDown(K As Integer, Shift As Integer)
Dim REC_SET As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
    If CHECKTB("asm.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If

    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from ASM where ASM_idy like '" & Sincrement(Trim(ASTIDY)) & "'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        ASTIDY = Sincrement(Trim(ASTIDY))
        ASTNME = Trim(REC_SET("ASM_nme") & "")
        indidy = Trim(REC_SET!ind_idy)
        If K = 38 Then
            ptyidy.SetFocus
        Else
            astqty.SetFocus
        End If
    Else
        V_AST_LST
    End If
End If
End Sub

Private Sub ASTIDY_KEYPRESS(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub astqty_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub astqty_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    If indidy = "Y" Then
    If MOD1 = False Then
        MS.Rows = Val(astqty) + 1
        For A = 1 To MS.Rows - 1
            MS.TextMatrix(A, 0) = A
        Next
        MS.Col = 1
    End If
        MS.SetFocus
    Else
        save
    End If
ElseIf K = 38 Then
    ASTIDY.SetFocus
End If
End Sub

Private Sub astqty_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
If TXNIDY.Enabled = True Then
    TXNIDY.SetFocus
Else
    TXNDAT.SetFocus
End If
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_AST_INW.ActiveControl Is TextBox Then M_FRM_AST_INW.ActiveControl.Text = Trim(FUNKEY(K))
If K = 78 And s = 2 Then
    newrec
ElseIf K = 83 And s = 2 Then
    save
ElseIf K = 81 And s = 2 Then
    cancel
ElseIf K = 88 And s = 2 Then
    Unload Me
ElseIf K = 80 And s = 2 Then
    M_FRM_PRV.Tag = "AST_INW"
    If CHECKTB("prt.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If

    If check("PRT", "PRT_NME", "M_FRM_AST_INW") = True Then M_FRM_PRV.Check3.Value = 1
    M_FRM_PRV.Show 1
ElseIf K = 112 Then
    M_HELP
ElseIf K = 69 And s = 2 Then
    If Me.ActiveControl.name = "ptyidy" Then V_DLR_LST
    If Me.ActiveControl.name = "ASTIDY" Then V_AST_LST
    If Me.ActiveControl.name = "TXNIDY" Then TXNLST
End If
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If list4.Visible = True Then
        If Label20 = "Select Party" Then
            list4.Visible = False
            SSTab1.Tab = 0
            ptyidy.SetFocus
        ElseIf Label20.Caption = "Select Asset" Then
            list4.Visible = False
            SSTab1.Tab = 0
            ASTIDY.SetFocus
        Else
            list4.Visible = False
            SSTab1.Tab = 0
            TXNIDY.SetFocus
        End If
    ElseIf Me.ActiveControl.name = "TXNIDY" Then
        Unload Me
    Else
        save
    End If
End If
End Sub

Private Sub Form_Load()
SETYEAR M_FRM_AST_INW
MS.ColWidth(0) = 500
MS.ColWidth(1) = MS.Width - 800
MS.ColAlignment(1) = 1
MS.TextMatrix(0, 1) = "Asset Code"
End Sub


Private Sub list4_DblClick()
list4_KeyPress 13
End Sub

Private Sub list4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If list4.ListItems.count > 0 Then
        If Label20 = "Select Party" Then
            ptyidy = list4.SelectedItem.ListSubItems.Item(1)
            PTYNME = list4.SelectedItem
            list4.Visible = False
            SSTab1.Tab = 0
            ASTIDY.SetFocus
        ElseIf Label20.Caption = "Select Asset" Then
            ASTIDY = list4.SelectedItem.ListSubItems.Item(1)
            ASTNME = list4.SelectedItem
            indidy = list4.SelectedItem.ListSubItems.Item(3)
            list4.Visible = False
            SSTab1.Tab = 0
            astqty.SetFocus
        Else
            TXNIDY = list4.SelectedItem.ListSubItems.Item(1)
            DISPDETAILS
            list4.Visible = False
            SSTab1.Tab = 0
            TXNDAT.SetFocus
        End If
    End If
End If
End Sub

Private Sub MS_KeyPress(K As Integer)
If MS.Col = 1 And MS.Row <> 0 Then
    If K = 13 Then
        If Len(Trim(MS.Text)) > 0 Then If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
    Else
        If K <> 27 Then
            Text3.Visible = True
            Text3.MaxLength = 10
            Text3.Top = MS.Top + MS.CellTop
            Text3.Left = MS.CellLeft + MS.Left
            Text3.Width = MS.CellWidth
            Text3.Height = MS.CellHeight
            Text3.Visible = True
            Text3.SetFocus
            If K <> 13 Then Text3 = Chr(K)
            Text3.SelStart = Len(Text3)
        End If
    End If
End If
End Sub

Private Sub MS_Scroll()
If Text3.Visible = True Then
    Text3 = ""
    Text3.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub ptyidy_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub PTYIDY_keydown(K As Integer, s As Integer)
Dim REC_SET As New ADODB.Recordset
If K = 13 Or K = 40 Or K = 38 Then
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

    If REC_SET.State = 1 Then REC_SET.Close
    REC_SET.Open "select * from act where act_idy like '" & Sincrement(Trim(ptyidy)) & "' AND (GRP_CLF LIKE 'Creditor' OR GRP_CLF LIKE 'Debtor') AND DBT_TYP='Y'", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
        M_CST_IDY = Sincrement(Trim(PTY_IDY))
        M_CST_NME = Trim(REC_SET("act_nme") & "")
        If K = 40 Or K = 13 Then
            ASTIDY.SetFocus
        Else
            TXNDAT.SetFocus
        End If
    Else
        V_DLR_LST
    End If
End If
End Sub
Private Sub V_DLR_LST()
If CHECKTB("act.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from act where (grp_clf LIKE 'Creditor' OR grp_clf LIKE 'Debtor') and act_nme like '" & Trim(ptyidy) & "%' AND DBT_TYP='Y'", con, adOpenKeyset, adLockPessimistic
list4.ListItems.CLEAR
list4.ColumnHeaders.CLEAR
list4.ColumnHeaders.Add , , "Name"
list4.ColumnHeaders.Add , , "ID"
list4.ColumnHeaders.Add , , "Address"
list4.ColumnHeaders.Item(1).Width = 3000
list4.ColumnHeaders.Item(2).Width = 1500
list4.ColumnHeaders.Item(3).Width = 2000
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("ACT_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_idy"))
        ls.ListSubItems.Add , , Trim(REC_SET("ACT_ad4") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Party"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Party Not Found, Define Party !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        ptyidy.SetFocus
    End If
End If
End Sub
Private Sub V_AST_LST()
If CHECKTB("asm.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim REC_SET As New ADODB.Recordset
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "select * from aSM WHERE aSM_nme like '" & Trim(ASTIDY) & "%'", con, adOpenKeyset, adLockPessimistic
list4.ListItems.CLEAR
list4.ColumnHeaders.CLEAR
list4.ColumnHeaders.Add , , "Name"
list4.ColumnHeaders.Add , , "ID"
list4.ColumnHeaders.Add , , "Quantity"
list4.ColumnHeaders.Add , , , 0
list4.ColumnHeaders.Item(1).Width = 2500
list4.ColumnHeaders.Item(2).Width = 1500
If REC_SET.RecordCount > 0 Then
    For A = 1 To REC_SET.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(REC_SET("ASM_nme")))
        ls.ListSubItems.Add , , Trim(REC_SET("ASM_idy"))
        ls.ListSubItems.Add , , Val(REC_SET("ASM_qty") & "")
        ls.ListSubItems.Add , , Trim(REC_SET("IND_IDY") & "")
        REC_SET.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Asset"
    list4.Visible = True
    list4.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Asset Not Found, Define Asset !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        ASTIDY.SetFocus
    End If
End If
End Sub

Private Sub ptyidy_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    If TXNIDY.Enabled = True Then TXNIDY.SetFocus
    If TXNIDY.Enabled = False Then TXNDAT.SetFocus
End If
End Sub

Private Sub text3_KeyPress(K As Integer)
If K = 13 Then
    For A = 1 To MS.Rows - 1
        If Trim(Text3) = Trim(MS.TextMatrix(A, 1)) Then
            MsgBox "ID already exists in this voucher", vbExclamation
            MS.Row = A
            MS.SetFocus
            Text3.Visible = False
            Exit Sub
        End If
    Next
    MS.Text = Text3
    Text3.Visible = False
    If MS.Row <> MS.Rows - 1 Then MS.Row = MS.Row + 1
    MS.SetFocus
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
ElseIf Button.KEY = "p" Then
    M_FRM_PRV.Tag = "AST_INW"
    If CHECKTB("prt.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If
    If check("PRT", "PRT_NME", "M_FRM_AST_INW") = True Then M_FRM_PRV.Check3.Value = 1
    M_FRM_PRV.Show 1
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If Me.ActiveControl.name = "ptyidy" Then V_DLR_LST
    If Me.ActiveControl.name = "ASTIDY" Then V_AST_LST
    If Me.ActiveControl.name = "TXNIDY" Then TXNLST
End If
End Sub

Public Sub newrec()
MOD1 = False
MS.Rows = 1
TXNIDY.Enabled = False
TXNIDY = ""
TXNDAT = Date
PTYNME = ""
ptyidy = ""
ASTIDY = ""
ASTNME = ""
astqty = 0
indidy = ""
TXNDAT.SetFocus
End Sub
Public Sub cancel()
MOD1 = False
MS.Rows = 1
TXNIDY.Enabled = True
TXNIDY = ""
TXNDAT = Date
PTYNME = ""
ptyidy = ""
ASTIDY = ""
ASTNME = ""
astqty = 0
indidy = ""
SSTab1.Tab = 0
TXNIDY.SetFocus
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.Tag = 1 Then
    SAL_QRY = "AST->TXN_TYP = 'AIN'"
    OpenReport CURDIR & "\Assets.RP1", "Asset Inward Entry", 2, SAL_QRY, 6
End If
End Sub

Private Sub TXNDAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    ptyidy.SetFocus
End If
End Sub

Private Sub TXNIDY_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(TXNIDY)) <> 0 Then
        DISPDETAILS
    Else
        TXNLST
    End If
ElseIf KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub TXNIDY_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Public Sub save()
If MsgBox("Confirm Transaction Save?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
If CHECKTB("act.dbf,asm.dbf,ast.dbf,prt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

If MOD1 = True Then
    INC = Trim(TXNIDY)
    MsgBox "Changes Not Saved", vbExclamation
    cancel
    GoTo n
End If
If check("ACT", "ACT_IDY", Trim(ptyidy)) = False Then
    MsgBox "Party Not Found, Define Party", vbExclamation
    ptyidy.SetFocus
    Exit Sub
End If
If check("Asm", "Asm_IDY", Trim(ASTIDY)) = False Then
    MsgBox "Asset Not Found, Define Asset", vbExclamation
    ASTIDY.SetFocus
    Exit Sub
End If
Dim RSINW As New ADODB.Recordset
Dim RSOUT As New ADODB.Recordset
RSINW.Open "select count(*),asm_cod,txn_typ from ast where asm_idy like '" & Trim(ASTIDY) & "' GROUP by ASM_COD,txn_typ", con, adOpenStatic, adLockReadOnly
Set RSOUT = RSINW.Clone
RSINW.filter = "txn_typ='AIN'"
RSOUT.filter = "txn_typ='AOU'"
For A = 1 To MS.Rows - 1
    If RSOUT.RecordCount > 0 Then RSOUT.MoveFirst
    RSOUT.Find "ASM_COD='" & Trim(MS.TextMatrix(A, 1)) & "'"
    If RSOUT.EOF = False Then
        AOUCOUNT = Val(RSOUT(0) & "")
        AINCOUNT = Val(RSINW(0) & "")
        If Val(AINCOUNT) <> Val(AOUCOUNT) Then
            MsgBox "Asset Code Already there", vbExclamation
            MS.Row = A
            MS.SetFocus
            Exit Sub
        End If
    Else
        If RSINW.RecordCount > 0 Then RSINW.MoveFirst
        RSINW.Find "ASM_COD='" & Trim(MS.TextMatrix(A, 1)) & "'"
        If RSINW.EOF = False Then
            MsgBox "Asset Code Already there", vbExclamation
            MS.Row = A
            MS.SetFocus
            Exit Sub
        End If
    End If
Next

If indidy = "Y" Then
    For A = 1 To MS.Rows - 1
        If Len(Trim(MS.TextMatrix(A, 1))) = 0 Then
            MsgBox "Invalid ID, Give ID", vbExclamation
            MS.Row = A
            MS.SetFocus
            Exit Sub
        End If
    Next
End If
If rs.State = 1 Then rs.Close
rs.Open "select * from ast", con, adOpenKeyset, adLockPessimistic
INC = "AIN" & T7increment("IDY_AIN")
If indidy = "Y" Then
    For A = 1 To MS.Rows - 1
        rs.AddNew
        rs("TXN_IDY") = INC
        rs("TXN_DAT") = Format(TXNDAT, "MM/DD/YYYY")
        rs("PTY_IDY") = Trim(ptyidy)
        rs("ASM_IDY") = Trim(ASTIDY)
        rs("ASM_COD") = Trim(MS.TextMatrix(A, 1))
        rs("AST_QTY") = 1
        rs!chk_flg = "Y"
        rs!TXN_TYP = "AIN"
        rs!AST_IDY = ""
        rs.Update
    Next
Else
    rs.AddNew
    rs("TXN_IDY") = INC
    INC = rs!TXN_IDY
    rs("TXN_DAT") = Format(TXNDAT, "MM/DD/YYYY")
    rs("PTY_IDY") = Trim(ptyidy)
    rs("ASM_IDY") = Trim(ASTIDY)
    rs("ASM_COD") = ""
    rs("AST_QTY") = Val(astqty)
    rs!TXN_TYP = "AIN"
    rs!chk_flg = "N"
    rs!AST_IDY = ""
    rs.Update
End If
If rs!chk_flg = "N" Then
    con.Execute "UPDATE ASM SET ASM_QTY = ASM_QTY + " & Val(astqty) & "WHERE ASM_IDY LIKE '" & Trim(ASTIDY) & "'"
Else
    con.Execute "UPDATE ASM SET ASM_QTY = ASM_QTY + " & Val(MS.Rows - 1) & "WHERE ASM_IDY LIKE '" & Trim(ASTIDY) & "'"
End If
MsgBox "New Transaction ID " & INC, vbExclamation
n:
cancel
If check("PRT", "PRT_NME", "M_FRM_AST_INW") = True Then OpenDocument CURDIR & "\VOUCH.RP1", "New Assets Inward Voucher", 2, Trim(INC), Trim(INC), 6, CURDIR & "\AST.cdx,c,TXN_IDY"
End Sub

Public Sub DISPDETAILS()
If CHECKTB("ast.dbf,act.dbf,asm.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset, RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Len(Trim(TXNIDY)) <> 10 Then
    rs.Open "select * from AST where txn_idy like 'AIN" & Tincrement(TXNIDY) & "' and txn_typ like 'AIN'", con, adOpenStatic
Else
    rs.Open "select * from AST where txn_idy like '" & Trim(TXNIDY) & "' and txn_typ like 'AIN'", con, adOpenStatic
End If
If rs.RecordCount > 0 Then
    If Len(Trim(TXNIDY)) <> 10 Then TXNIDY = "AIN" & Tincrement(TXNIDY)
    MOD1 = True
    TXNDAT = Trim(rs("txn_dat"))
    ptyidy = Trim(rs!PTY_IDY)
    PTYNME = GetValue("ACT_NME", "ACT WHERE ACT_IDY LIKE '" & Trim(ptyidy) & "'")
    ASTIDY = Trim(rs!ASM_IDY)
    ASTNME = GetValue("ASM_NME", "ASM WHERE ASM_IDY LIKE '" & Trim(ASTIDY) & "'")
    If rs!chk_flg <> "Y" Then
        astqty = Trim(rs!AST_QTY)
    Else
        astqty = Val(rs.RecordCount)
    End If
    TXNIDY.Enabled = False
    SSTab1.Tab = 0
    TXNDAT.SetFocus
    If rs!chk_flg = "Y" Then
        indidy = rs!chk_flg
        MS.Rows = rs.RecordCount + 1
        For A = 1 To rs.RecordCount
            MS.TextMatrix(A, 0) = A
            MS.TextMatrix(A, 1) = Trim(rs("ASM_COD") & "")
            rs.MoveNext
        Next
    End If
Else
    MsgBox "Asset Inwards Voucher Number Not Found !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    TXNIDY.Enabled = True
    TXNIDY.SetFocus
End If
Exit Sub
LBL:
If err.Number = 35773 Then
    MsgBox "Invalid Asset Voucher Selected for this Financial Year", vbExclamation
    cancel
Else
    MsgBox err.Description
    cancel
End If
End Sub
Public Sub TXNLST()
If CHECKTB("ast.dbf,act.dbf,asm.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select ast.txn_idy,ast.txn_dat,act_nme,asm_nme from ast,act,asm where ast.txn_typ like 'AIN' and act.act_idy like AST.PTY_idy AND AST.ASM_IDY LIKE ASM.ASM_IDY GROUP BY TXN_IDY", con, adOpenKeyset, adLockPessimistic
list4.ListItems.CLEAR
list4.ColumnHeaders.CLEAR
list4.ColumnHeaders.Add , , "Asset Name", 2500
list4.ColumnHeaders.Add , , "ID", 1500
list4.ColumnHeaders.Add , , "Date", 1500
list4.ColumnHeaders.Add , , "Party Name", 2500
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list4.ListItems.Add(, , Trim(rs(3)))
        ls.ListSubItems.Add , , Trim(rs(0))
        ls.ListSubItems.Add , , Format(rs(1), "dd/mmm/yyyy")
        ls.ListSubItems.Add , , Trim(rs(2))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Asset Voucher"
    list4.Visible = True
    list4.SetFocus
Else
    MsgBox "Asset Voucher Not Found !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    TXNIDY.Enabled = True
    TXNIDY.SetFocus
End If
End Sub

Public Function CHECKASSETS() As Boolean

End Function
