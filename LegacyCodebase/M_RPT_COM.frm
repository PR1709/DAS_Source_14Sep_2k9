VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_RPT_COM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analysis Combination Reports"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   Icon            =   "M_RPT_COM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5145
      Left            =   -15
      TabIndex        =   1
      Top             =   480
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Report Combination &Filters"
      TabPicture(0)   =   "M_RPT_COM.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TDATE"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FDATE"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lv1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbrowlist"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmbcollist"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbCaldates"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Check8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkPen"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      Begin VB.CheckBox chkPen 
         Caption         =   "&Penetration"
         Height          =   255
         Left            =   7320
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox Check8 
         Caption         =   "&FMCG Calendar"
         Height          =   375
         Left            =   150
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cmbCaldates 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   13
         Top             =   540
         Width           =   615
      End
      Begin VB.ComboBox cmbcollist 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1620
         Width           =   2415
      End
      Begin VB.ComboBox cmbrowlist 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1140
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   615
         Left            =   4800
         TabIndex        =   4
         Top             =   405
         Width           =   2415
         Begin VB.OptionButton Option2 
            Caption         =   "&Retailer "
            Height          =   255
            Left            =   1080
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&SKU "
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView lv1 
         Height          =   3030
         Left            =   45
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2055
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   5345
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
      Begin MSComCtl2.DTPicker FDATE 
         Height          =   330
         Left            =   1680
         TabIndex        =   16
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16515075
         CurrentDate     =   38031
      End
      Begin MSComCtl2.DTPicker TDATE 
         Height          =   330
         Left            =   1680
         TabIndex        =   17
         Top             =   1575
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16515075
         CurrentDate     =   38031
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Report Columns"
         Height          =   195
         Left            =   3570
         TabIndex        =   10
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Report Rows"
         Height          =   195
         Left            =   3720
         TabIndex        =   8
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Base Selection"
         Height          =   195
         Left            =   3600
         TabIndex        =   5
         Top             =   615
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   1335
         TabIndex        =   3
         Top             =   1643
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   1118
         Width           =   345
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
            Picture         =   "M_RPT_COM.frx":045E
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_COM.frx":09A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_COM.frx":0D6A
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_COM.frx":10BE
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_COM.frx":16F6
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_COM.frx":1A4A
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_COM.frx":1EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_COM.frx":219E
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_COM.frx":25AA
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_COM.frx":26B6
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_COM.frx":2A0A
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_COM.frx":2E1E
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
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
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "n"
            Object.ToolTipText     =   "New (Ctrl + N)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s"
            Object.ToolTipText     =   "Save (Ctrl + S)"
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
            Object.Visible         =   0   'False
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sku"
                  Object.Tag             =   "sku"
                  Text            =   "SKU Wise Route Sales"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_RPT_COM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsprd As New ADODB.Recordset
Dim rsret As New ADODB.Recordset
Dim RSCAT As New ADODB.Recordset
Dim rsloc As New ADODB.Recordset
Dim rsmar As New ADODB.Recordset
Dim RSSMN As New ADODB.Recordset
Dim rsrou As New ADODB.Recordset
Dim rsdir As New ADODB.Recordset
Dim RSWEK As New ADODB.Recordset
Dim n As Integer
Dim ARR As String, ss As String
Dim CNT As Integer, didy As String, didx As String

Private Sub chkPen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmbcollist_KeyDown 13, 0
    CNT = 0
    lv1.SetFocus
End If
End Sub

Private Sub cmbCaldates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If RSWEK.State = 1 Then RSWEK.Close
RSWEK.Open "select frm_dat,to_dat from wek where wek_idy=" & Val(cmbCaldates.Text), con, adOpenKeyset, adLockPessimistic
If Not RSWEK.EOF Then
    FDATE.Enabled = True
    TDATE.Enabled = True
    FDATE.Value = RSWEK("frm_dat")
    TDATE.Value = RSWEK("to_dat")
    FDATE.Enabled = False
    TDATE.Enabled = False
    Frame1.Enabled = True
    Option1.SetFocus
End If
End If
End Sub

Private Sub cmbcollist_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CNT = 0
    lv1.ListItems.CLEAR
    If Len(Trim(cmbcollist.Text)) > 0 Then
        If Trim(cmbcollist.Text) = "SKU Name" Then
            LVP
            PRDLIST
            didy = "PRD"
            lv1.SetFocus
        ElseIf Trim(cmbcollist.Text) = "Pack Size" Then
            With lv1
                .ColumnHeaders.CLEAR
                .ColumnHeaders.Add , , "Pack Size", 2500
'                .ColumnHeaders.Add , , "Product ID", 1500
            End With
            packlist
            didy = "BRD"
            lv1.SetFocus
        ElseIf Trim(cmbcollist.Text) = "Pack Type" Then
            With lv1
                .ColumnHeaders.CLEAR
                .ColumnHeaders.Add , , "Pack Type", 2500
 '               .ColumnHeaders.Add , , "Product ID", 1500
            End With
            ptypelist
            didy = "BRD"
            lv1.SetFocus
        ElseIf Trim(cmbcollist.Text) = "Brand" Then
            With lv1
                .ColumnHeaders.CLEAR
                .ColumnHeaders.Add , , "Brand Name", 2500
  '              .ColumnHeaders.Add , , "Product ID", 1500
            End With
            brdlist
            didy = "BRD"
            lv1.SetFocus
        ElseIf Trim(cmbcollist.Text) = "Segment" Then
            With lv1
                .ColumnHeaders.CLEAR
                .ColumnHeaders.Add , , "Segment", 2500
   '             .ColumnHeaders.Add , , "Group ID", 2500
            End With
            seglist
            didy = "SEG"
            lv1.SetFocus
        ElseIf Trim(cmbcollist.Text) = "Category" Then
            With lv1
                .ColumnHeaders.CLEAR
                .ColumnHeaders.Add , , "Category", 2500
    '            .ColumnHeaders.Add , , "Group ID", 2500
            End With
            catlist
            didy = "PGR"
            lv1.SetFocus
        ElseIf Trim(cmbcollist.Text) = "Retailer" Then
            With lv1
                .ColumnHeaders.CLEAR
                .ColumnHeaders.Add , , "Retailer Name", 2500
                .ColumnHeaders.Add , , "ID", 1500
                .ColumnHeaders.Add , , "Market", 2500
            End With
            retlist
            didy = "DLR"
            lv1.SetFocus
        ElseIf Trim(cmbcollist.Text) = "Ret.Category" Then
            With lv1
                .ColumnHeaders.CLEAR
                .ColumnHeaders.Add , , "Category Name", 2500
                .ColumnHeaders.Add , , "ID", 1500
            End With
            rcatlist
            didy = "CAT"
            lv1.SetFocus
        ElseIf Trim(cmbcollist.Text) = "Location" Then
            With lv1
                .ColumnHeaders.CLEAR
                .ColumnHeaders.Add , , "Location Name", 2500
                .ColumnHeaders.Add , , "ID", 1500
            End With
            loclist
            didy = "LOC"
            lv1.SetFocus
        ElseIf Trim(cmbcollist.Text) = "Market" Then
            With lv1
                .ColumnHeaders.CLEAR
                .ColumnHeaders.Add , , "Market Name", 2500
                .ColumnHeaders.Add , , "ID", 1500
            End With
            marlist
            didy = "MAR"
            lv1.SetFocus
        ElseIf Trim(cmbcollist.Text) = "Salesman" Then
            With lv1
                .ColumnHeaders.CLEAR
                .ColumnHeaders.Add , , "Salesman Name", 2500
                .ColumnHeaders.Add , , "ID", 1500
            End With
            smnlist
            didy = "SMN"
            lv1.SetFocus
        ElseIf Trim(cmbcollist.Text) = "Route" Then
            With lv1
                .ColumnHeaders.CLEAR
                .ColumnHeaders.Add , , "Route ID", 1500
            End With
            roulist
            didy = "ROU"
            lv1.SetFocus
        End If
    Else
        MsgBox "Select Column Value", vbExclamation, "DAS Version 4.0"
        cmbcollist.SetFocus
    End If
End If
End Sub

Private Sub cmbcollist_LostFocus()
If Trim(cmbrowlist.Text) = "Route" And Trim(cmbcollist.Text) = "Brand" Then
   chkPen.Enabled = True
   chkPen.SetFocus
ElseIf Trim(cmbrowlist.Text) = "Route" And Trim(cmbcollist.Text) = "SKU Name" Then
   chkPen.Enabled = True
   chkPen.SetFocus
Else
   chkPen.Enabled = False
End If
End Sub

Private Sub cmbrowlist_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Trim(cmbrowlist.Text) = "" Then
        MsgBox "Select Row Value", vbExclamation, "DAS Version 4.0"
        cmbrowlist.SetFocus
    Else
        cmbcollist.SetFocus
    End If
End If
End Sub

Private Sub FDATE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    TDATE.SetFocus
End If

End Sub

Private Sub Form_Activate()
SETYEAROPEN M_RPT_COM
Check8.SetFocus
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf Me.ActiveControl Is TextBox Then Me.ActiveControl.Text = Trim(FUNKEY(K))
If K = 78 And s = 2 Then
    newrecord
ElseIf K = 83 And s = 2 Then
    saverecord
ElseIf K = 81 And s = 2 Then
    cancel
ElseIf K = 88 And s = 2 Then
    Unload Me
ElseIf K = 112 Then
    M_HELP
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If Me.ActiveControl.name = "Check8" Then
        Unload Me
    ElseIf Me.ActiveControl.name = "lv1" Then
        If lv1.ListItems.count >= 1 Then
            saverecord
        End If
    Else
        saverecord
    End If
End If
End Sub

Private Sub Form_Load()
SETYEAROPEN Me
chkPen.Enabled = False
FDATE.Value = Now
TDATE.Value = Now
lv1.View = lvwReport
CNT = 0
didx = "IDY"
End Sub

Private Sub lv1_Click()
CNT = 0
For i = 1 To lv1.ListItems.count
    lv1.ListItems(i).Checked = False
Next
End Sub

Private Sub lv1_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then
    If lv1.ListItems(lv1.SelectedItem.Index).Checked = False Then
        CNT = CNT - 1
    Else
        If CNT < 5 Then
            CNT = CNT + 1
        Else
            MsgBox "Only 5 Entities Can Be Selected As Columns for Report!", vbExclamation
            lv1.ListItems(lv1.SelectedItem.Index).Checked = False
            lv1.SetFocus
        End If
    End If
End If
If KeyAscii = 13 Then
    saverecord
End If
End Sub


Private Sub Option1_Click()
addskulist cmbrowlist
addretlist cmbcollist
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmbrowlist.SetFocus
End Sub

Private Sub Option2_Click()
addretlist cmbrowlist
addskulist cmbcollist
End Sub

Private Sub addskulist(cm As ComboBox)
With cm
    .CLEAR
    .AddItem "SKU Name"
    .AddItem "Brand"
    .AddItem "Segment"
    .AddItem "Category"
    .AddItem "Pack Type"
    .AddItem "Pack Size"
End With
End Sub

Private Sub addretlist(cm As ComboBox)
With cm
    .CLEAR
    .AddItem "Retailer"
    .AddItem "Ret.Category"
    .AddItem "Location"
    .AddItem "Market"
    .AddItem "Salesman"
    .AddItem "Route"
End With
End Sub

Private Sub PRDLIST()
If rsprd.State = 1 Then rsprd.Close
rsprd.Open "SELECT PRD.PRD_IDY,PRD.PRD_NME,PRD.PRD_PCK,PRD.GRP_IDY,PGR.GRP_CLS AS PRD_SEG FROM PRD,PGR WHERE PRD.GRP_IDY LIKE PGR.GRP_IDY", con, adOpenKeyset, adLockPessimistic
For i = 0 To rsprd.RecordCount - 1
    Set ls = lv1.ListItems.Add(, , rsprd("PRD_NME"))
    ls.ListSubItems.Add , , rsprd("PRD_IDY")
    ls.ListSubItems.Add , , rsprd("PRD_PCK")
    ls.ListSubItems.Add , , rsprd("PRD_SEG")
    rsprd.MoveNext
Next
End Sub

Private Sub packlist()
Dim ls As ListItem
If rsprd.State = 1 Then rsprd.Close
rsprd.Open "select prd_pck from prd GROUP BY PRD_PCK", con, adOpenKeyset, adLockPessimistic
lv1.ListItems.CLEAR
For i = 0 To rsprd.RecordCount - 1
    If Len(Trim(rsprd("prd_pck"))) > 0 Then
        Set ls = lv1.ListItems.Add(, , rsprd("prd_pck"))
    End If
    rsprd.MoveNext
Next
End Sub

Private Sub ptypelist()
If rsprd.State = 1 Then rsprd.Close
Dim ls As ListItem
rsprd.Open "select prd_pkt from prd GROUP BY PRD_PKT", con, adOpenKeyset, adLockPessimistic
lv1.ListItems.CLEAR
For i = 0 To rsprd.RecordCount - 1
    If Len(Trim(rsprd("prd_pkt"))) > 0 Then
        Set ls = lv1.ListItems.Add(, , rsprd("prd_pkt"))
    End If
    rsprd.MoveNext
Next
End Sub

Private Sub brdlist()
If rsprd.State = 1 Then rsprd.Close
Dim ls As ListItem
lv1.ListItems.CLEAR
rsprd.Open "select brd_nme from prd group by brd_nme", con, adOpenKeyset, adLockPessimistic
For i = 0 To rsprd.RecordCount - 1
    If Len(Trim(rsprd("brd_nme"))) > 0 Then
        Set ls = lv1.ListItems.Add(, , rsprd("brd_nme"))
    End If
    rsprd.MoveNext
Next
End Sub

Private Sub seglist()
If rsprd.State = 1 Then rsprd.Close
rsprd.Open "select grp_cls from pgr group by grp_cls", con, adOpenKeyset, adLockPessimistic
lv1.ListItems.CLEAR
For i = 0 To rsprd.RecordCount - 1
    If Len(Trim(rsprd("grp_cls"))) > 0 Then
        Set ls = lv1.ListItems.Add(, , rsprd("grp_cls"))
    End If
    rsprd.MoveNext
Next
End Sub
Private Sub catlist()
Dim ls As ListItem
If rsprd.State = 1 Then rsprd.Close
rsprd.Open "select grp_cat from pgr GROUP BY GRP_CAT", con, adOpenKeyset, adLockPessimistic
lv1.ListItems.CLEAR
For i = 0 To rsprd.RecordCount - 1
    If Len(Trim(rsprd("grp_cat"))) > 0 Then
        Set ls = lv1.ListItems.Add(, , rsprd("grp_cat"))
    End If
    rsprd.MoveNext
Next
End Sub

Private Sub LVP()
With lv1
    .ColumnHeaders.CLEAR
    .ColumnHeaders.Add , , "Product Name", 3500
    .ColumnHeaders.Add , , "Product ID", 1700
    .ColumnHeaders.Add , , "Pack Size", 1700
    .ColumnHeaders.Add , , "Segment", 1800
End With
End Sub

Private Sub cancel()
chkPen.Enabled = False
cmbcollist.CLEAR
FDATE.Enabled = True
TDATE.Enabled = True
FDATE.Value = Date
TDATE.Value = Date
Option1.Value = False
Option2.Value = False
cmbrowlist.CLEAR
cmbcollist.CLEAR
lv1.ListItems.CLEAR
cmbCaldates.CLEAR
Check8.Value = 0
Check8.SetFocus
chkPen.Value = 0
CNT = 0
End Sub
Private Sub retlist()
Dim ls As ListItem
If rsret.State = 1 Then rsret.Close
rsret.Open "select act_idy,act_nme,mar.* from act,mar where grp_clf='Debtor' and act.mar_idy=mar.mar_idy", con, adOpenKeyset, adLockPessimistic
With lv1
    .ListItems.CLEAR
    While Not rsret.EOF
        Set ls = .ListItems.Add(, , rsret("act_nme"))
        ls.ListSubItems.Add , , rsret("act_idy")
        ls.ListSubItems.Add , , rsret!mar_nme
        rsret.MoveNext
    Wend
End With
End Sub
Private Sub rcatlist()
Dim ls As ListItem
If RSCAT.State = 1 Then RSCAT.Close
RSCAT.Open "select cat_idy,cat_nme from cat", con, adOpenKeyset, adLockPessimistic
With lv1
    .ListItems.CLEAR
    While Not RSCAT.EOF
        Set ls = .ListItems.Add(, , RSCAT("cat_nme"))
        ls.ListSubItems.Add , , RSCAT("cat_idy")
        RSCAT.MoveNext
    Wend
End With
End Sub
Private Sub loclist()
Dim ls As ListItem
If rsloc.State = 1 Then rsloc.Close
rsloc.Open "Select loc_idy,loc_nme from loc", con, adOpenKeyset, adLockPessimistic
With lv1
    .ListItems.CLEAR
    While Not rsloc.EOF
        Set ls = .ListItems.Add(, , rsloc("loc_nme"))
        ls.ListSubItems.Add , , rsloc("loc_idy")
        rsloc.MoveNext
    Wend
End With
End Sub
Private Sub marlist()
Dim ls As ListItem
If rsmar.State = 1 Then rsmar.Close
rsmar.Open "Select mar_idy,mar_nme from mar", con, adOpenKeyset, adLockPessimistic
With lv1
    .ListItems.CLEAR
    While Not rsmar.EOF
        Set ls = .ListItems.Add(, , rsmar("mar_nme"))
        ls.ListSubItems.Add , , rsmar("mar_idy")
        rsmar.MoveNext
    Wend
End With
End Sub

Private Sub smnlist()
Dim ls As ListItem
If RSSMN.State = 1 Then RSSMN.Close
RSSMN.Open "Select smn_idy,smn_nme from smn", con, adOpenKeyset, adLockPessimistic
With lv1
    .ListItems.CLEAR
    While Not RSSMN.EOF
        Set ls = .ListItems.Add(, , RSSMN("smn_nme"))
        ls.ListSubItems.Add , , RSSMN("smn_idy")
        RSSMN.MoveNext
    Wend
End With
End Sub
Private Sub roulist()
Dim ls As ListItem
If rsrou.State = 1 Then rsrou.Close
rsrou.Open "select rou_idy from rdy group by rou_idy", con, adOpenKeyset, adLockPessimistic
With lv1
    .ListItems.CLEAR
    While Not rsrou.EOF
        Set ls = .ListItems.Add(, , Trim(rsrou!rou_idy & ""))
        rsrou.MoveNext
    Wend
End With
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmbrowlist.SetFocus
End If
End Sub

Private Sub reptyp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If reptyp.ListIndex = 0 Then
        Frame1.Enabled = True
        Option1.SetFocus
    ElseIf reptyp.ListIndex = 1 Then
        ' Growth rate
    ElseIf reptyp.ListIndex = 2 Then
        ' Penetration reports
    End If
End If
End Sub

Private Sub TDATE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    TDATE.Refresh
    If TDATE.Value >= FDATE.Value Then
        Frame1.Enabled = True
        Option1.SetFocus
    Else
        MsgBox "Invalid Date", vbExclamation, "DAS Version 4.0"
        TDATE.SetFocus
    End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    newrecord
ElseIf Button.KEY = "s" Then
    saverecord
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "p" Then

ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "e" Then
    Unload Me
End If
End Sub
Private Sub newrecord()
    cancel
End Sub
Private Sub saverecord()
If MsgBox("Save Selection for Report?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
If Trim(cmbrowlist.Text) = "" Then
    MsgBox "Row Selection Required!", vbExclamation, "DAS Version 4.0"
    cmbrowlist.SetFocus
    Exit Sub
End If
If Trim(cmbcollist.Text) = "" Then
    MsgBox "Column Selection Required!", vbExclamation, "DAS Version 4.0"
    cmbcollist.SetFocus
    Exit Sub
End If
ss = CALLIDS()
If n = 0 Then
    MsgBox "Selection Required for Report Columns!", vbExclamation, "DAS Version 4.0"
    lv1.SetFocus
    Exit Sub
End If
con.Execute "Update dir set dir_msg='" & ss & "' where dir_idy='" & didy & "' and dir_idx='" & didx & "'"
If rsdir.State = 1 Then rsdir.Close
rsdir.Open "select * from dir where dir_idy='HDR' and dir_idx='DAT'", con, adOpenKeyset, adLockPessimistic
If Not rsdir.EOF Then
    rsdir("dir_als") = Format(FDATE.Value, "dd/MM/yyyy")
    rsdir("DIR_MSG") = Format(TDATE.Value, "dd/MM/yyyy")
    rsdir.Update
End If

If chkPen.Value = 1 Then
    If Trim(cmbrowlist.Text) = "Route" And Trim(cmbcollist.Text) = "Brand" Then
        AddScheduledCalls FDATE, TDATE, "ROU"
        doPent FDATE, TDATE, "ROU"
        OpenReport CURDIR & "\PNT_MIS.RP1", "Route Wise Brand Penetration", 2, "", 6
    ElseIf Trim(cmbrowlist.Text) = "Route" And Trim(cmbcollist.Text) = "SKU Name" Then
        AddScheduledCalls FDATE, TDATE, "ROU"
        doPentSKU FDATE, TDATE, "ROU"
        OpenReport CURDIR & "\PNT_MIS.RP1", "Route Wise SKU Penetration", 2, "", 6
    Else
        MsgBox "Invalid Combination for Penetration!", vbExclamation, "DAS Version 4.0"
    End If
    cancel
    Exit Sub
End If
SAL_QRY = "inV->inv_dat >= ctod('" & Format(FDATE.Value, "MM/DD/YYYY") & "').and.inv->inv_dat <= ctod('" & Format(TDATE.Value, "MM/DD/YYYY") & "').AND.INV->TXN_TYP='SAL'.AND.(INV->STK_TYP='Salable'.OR.INV->STK_TYP='Free')"
If Trim(cmbrowlist.Text) = "SKU Name" And Trim(cmbcollist.Text) = "Retailer" Then
    OpenReport CURDIR & "\SKU_MIS.RP1", "SKU Wise Retailer Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "SKU Name" And Trim(cmbcollist.Text) = "Ret.Category" Then
    OpenReport CURDIR & "\SKU_MIS.RP1", "SKU Wise Category Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "SKU Name" And Trim(cmbcollist.Text) = "Market" Then
    OpenReport CURDIR & "\SKU_MIS.RP1", "SKU Wise Market Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "SKU Name" And Trim(cmbcollist.Text) = "Location" Then
    OpenReport CURDIR & "\SKU_MIS.RP1", "SKU Wise Location Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "SKU Name" And Trim(cmbcollist.Text) = "Route" Then
    OpenReport CURDIR & "\SKU_MIS.RP1", "SKU Wise Route Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "SKU Name" And Trim(cmbcollist.Text) = "Salesman" Then
    OpenReport CURDIR & "\SKU_MIS.RP1", "SKU Wise Salesman Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Category" And Trim(cmbcollist.Text) = "Ret.Category" Then
    OpenReport CURDIR & "\CAT_MIS.RP1", "Cat Wise Category Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Category" And Trim(cmbcollist.Text) = "Location" Then
    OpenReport CURDIR & "\CAT_MIS.RP1", "Cat Wise Location Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Category" And Trim(cmbcollist.Text) = "Market" Then
    OpenReport CURDIR & "\CAT_MIS.RP1", "Cat Wise Market Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Category" And Trim(cmbcollist.Text) = "Retailer" Then
    OpenReport CURDIR & "\CAT_MIS.RP1", "Cat Wise Retailer Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Category" And Trim(cmbcollist.Text) = "Route" Then
    OpenReport CURDIR & "\CAT_MIS.RP1", "Cat Wise Route Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Category" And Trim(cmbcollist.Text) = "Salesman" Then
    OpenReport CURDIR & "\CAT_MIS.RP1", "Cat Wise Salesman Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Location" And Trim(cmbcollist.Text) = "Brand" Then
    OpenReport CURDIR & "\LOC_MIS.RP1", "Location Brand Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Location" And Trim(cmbcollist.Text) = "Category" Then
    OpenReport CURDIR & "\LOC_MIS.RP1", "Location Category Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Location" And Trim(cmbcollist.Text) = "Pack Size" Then
    OpenReport CURDIR & "\LOC_MIS.RP1", "Location Pack Size Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Location" And Trim(cmbcollist.Text) = "Pack Type" Then
    OpenReport CURDIR & "\LOC_MIS.RP1", "Location Pack Type Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Location" And Trim(cmbcollist.Text) = "Segment" Then
    OpenReport CURDIR & "\LOC_MIS.RP1", "Location Segment Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Location" And Trim(cmbcollist.Text) = "SKU Name" Then
    OpenReport CURDIR & "\LOC_MIS.RP1", "Location SKU Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Market" And Trim(cmbcollist.Text) = "Brand" Then
    OpenReport CURDIR & "\MAR_MIS.RP1", "Market Brand Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Market" And Trim(cmbcollist.Text) = "Category" Then
    OpenReport CURDIR & "\MAR_MIS.RP1", "Market Category Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Market" And Trim(cmbcollist.Text) = "Pack Size" Then
    OpenReport CURDIR & "\MAR_MIS.RP1", "Market Pack Size Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Market" And Trim(cmbcollist.Text) = "Pack Type" Then
    OpenReport CURDIR & "\MAR_MIS.RP1", "Market Pack Type Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Market" And Trim(cmbcollist.Text) = "Segment" Then
    OpenReport CURDIR & "\MAR_MIS.RP1", "Market Segment Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Market" And Trim(cmbcollist.Text) = "SKU Name" Then
    OpenReport CURDIR & "\MAR_MIS.RP1", "Market SKU Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Type" And Trim(cmbcollist.Text) = "Ret.Category" Then
    OpenReport CURDIR & "\PKT_MIS.RP1", "Pack Type Category Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Type" And Trim(cmbcollist.Text) = "Location" Then
    OpenReport CURDIR & "\PKT_MIS.RP1", "Pack Type Location Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Type" And Trim(cmbcollist.Text) = "Market" Then
    OpenReport CURDIR & "\PKT_MIS.RP1", "Pack Type Market Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Type" And Trim(cmbcollist.Text) = "Retailer" Then
    OpenReport CURDIR & "\PKT_MIS.RP1", "Pack Type Retailer Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Type" And Trim(cmbcollist.Text) = "Route" Then
    OpenReport CURDIR & "\PKT_MIS.RP1", "Pack Type Route Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Type" And Trim(cmbcollist.Text) = "Salesman" Then
    OpenReport CURDIR & "\PKT_MIS.RP1", "Pack Type Salesman Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Size" And Trim(cmbcollist.Text) = "Ret.Category" Then
    OpenReport CURDIR & "\PCK_MIS.RP1", "Pack Size Category Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Size" And Trim(cmbcollist.Text) = "Location" Then
    OpenReport CURDIR & "\PCK_MIS.RP1", "Pack Size Location Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Size" And Trim(cmbcollist.Text) = "Market" Then
    OpenReport CURDIR & "\PCK_MIS.RP1", "Pack Size Market Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Size" And Trim(cmbcollist.Text) = "Retailer" Then
    OpenReport CURDIR & "\PCK_MIS.RP1", "Pack Size Retailer Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Size" And Trim(cmbcollist.Text) = "Route" Then
    OpenReport CURDIR & "\PCK_MIS.RP1", "Pack Size Route Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Pack Size" And Trim(cmbcollist.Text) = "Salesman" Then
    OpenReport CURDIR & "\PCK_MIS.RP1", "Pack Size Salesman Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Retailer" And Trim(cmbcollist.Text) = "Brand" Then
    OpenReport CURDIR & "\RET_MIS.RP1", "Retailer Brand Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Retailer" And Trim(cmbcollist.Text) = "Category" Then
    OpenReport CURDIR & "\RET_MIS.RP1", "Retailer Category Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Retailer" And Trim(cmbcollist.Text) = "Pack Size" Then
    OpenReport CURDIR & "\RET_MIS.RP1", "Retailer Pack Size Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Retailer" And Trim(cmbcollist.Text) = "Pack Type" Then
    OpenReport CURDIR & "\RET_MIS.RP1", "Retailer Pack Type Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Retailer" And Trim(cmbcollist.Text) = "Segment" Then
    OpenReport CURDIR & "\RET_MIS.RP1", "Retailer Segment Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Retailer" And Trim(cmbcollist.Text) = "SKU Name" Then
    OpenReport CURDIR & "\RET_MIS.RP1", "Retailer SKU Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Route" And Trim(cmbcollist.Text) = "Brand" Then
    OpenReport CURDIR & "\ROU_MIS.RP1", "Route Wise Brand Sales Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Route" And Trim(cmbcollist.Text) = "Category" Then
    OpenReport CURDIR & "\ROU_MIS.RP1", "Route Wise Category Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Route" And Trim(cmbcollist.Text) = "Pack Size" Then
    OpenReport CURDIR & "\ROU_MIS.RP1", "Route Wise Pack Size Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Route" And Trim(cmbcollist.Text) = "Pack Type" Then
    OpenReport CURDIR & "\ROU_MIS.RP1", "Route Wise Pack Type Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Route" And Trim(cmbcollist.Text) = "Segment" Then
    OpenReport CURDIR & "\ROU_MIS.RP1", "Route Wise Segment Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Route" And Trim(cmbcollist.Text) = "SKU Name" Then
    OpenReport CURDIR & "\ROU_MIS.RP1", "Route Wise SKU Wise Report", 2, Trim(SAL_QRY), 6
 ElseIf Trim(cmbrowlist.Text) = "Ret.Category" And Trim(cmbcollist.Text) = "Brand" Then
    OpenReport CURDIR & "\RTC_MIS.RP1", "Category Brand Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Ret.Category" And Trim(cmbcollist.Text) = "Category" Then
    OpenReport CURDIR & "\RTC_MIS.RP1", "Category Category Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Ret.Category" And Trim(cmbcollist.Text) = "Pack Size" Then
    OpenReport CURDIR & "\RTC_MIS.RP1", "Category Pack Size Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Ret.Category" And Trim(cmbcollist.Text) = "Pack Type" Then
    OpenReport CURDIR & "\RTC_MIS.RP1", "Category Pack Type Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Ret.Category" And Trim(cmbcollist.Text) = "Segment" Then
    OpenReport CURDIR & "\RTC_MIS.RP1", "Category Segment Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Ret.Category" And Trim(cmbcollist.Text) = "SKU Name" Then
    OpenReport CURDIR & "\RTC_MIS.RP1", "Category SKU Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Segment" And Trim(cmbcollist.Text) = "Ret.Category" Then
    OpenReport CURDIR & "\SEG_MIS.RP1", "Seg Wise Category Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Segment" And Trim(cmbcollist.Text) = "Location" Then
    OpenReport CURDIR & "\SEG_MIS.RP1", "Seg Wise Location Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Segment" And Trim(cmbcollist.Text) = "Market" Then
    OpenReport CURDIR & "\SEG_MIS.RP1", "Seg Wise Market Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Segment" And Trim(cmbcollist.Text) = "Retailer" Then
    OpenReport CURDIR & "\SEG_MIS.RP1", "Seg Wise Retailer Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Segment" And Trim(cmbcollist.Text) = "Route" Then
    OpenReport CURDIR & "\SEG_MIS.RP1", "Seg Wise Route Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Segment" And Trim(cmbcollist.Text) = "Salesman" Then
    OpenReport CURDIR & "\SEG_MIS.RP1", "Seg Wise Salesman Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Salesman" And Trim(cmbcollist.Text) = "Brand" Then
    OpenReport CURDIR & "\SMN_MIS.RP1", "Salesman Brand Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Salesman" And Trim(cmbcollist.Text) = "Category" Then
    OpenReport CURDIR & "\SMN_MIS.RP1", "Salesman Category Report", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Salesman" And Trim(cmbcollist.Text) = "Pack Size" Then
    OpenReport CURDIR & "\SMN_MIS.RP1", "Salesman Pack Size Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Salesman" And Trim(cmbcollist.Text) = "Pack Type" Then
    OpenReport CURDIR & "\SMN_MIS.RP1", "Salesman Pack Type Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Salesman" And Trim(cmbcollist.Text) = "Segment" Then
    OpenReport CURDIR & "\SMN_MIS.RP1", "Salesman Segment Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Salesman" And Trim(cmbcollist.Text) = "SKU Name" Then
    OpenReport CURDIR & "\SMN_MIS.RP1", "Salesman SKU Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Brand" And Trim(cmbcollist.Text) = "Ret.Category" Then
    OpenReport CURDIR & "\BRD_MIS.RP1", "Brand Wise Category Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Brand" And Trim(cmbcollist.Text) = "Location" Then
    OpenReport CURDIR & "\BRD_MIS.RP1", "Brand Wise Location Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Brand" And Trim(cmbcollist.Text) = "Market" Then
    OpenReport CURDIR & "\BRD_MIS.RP1", "Brand Wise Market Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Brand" And Trim(cmbcollist.Text) = "Retailer" Then
    OpenReport CURDIR & "\BRD_MIS.RP1", "Brand Wise Retailer Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Brand" And Trim(cmbcollist.Text) = "Route" Then
    OpenReport CURDIR & "\BRD_MIS.RP1", "Brand Wise Route Wise Sales", 2, Trim(SAL_QRY), 6
ElseIf Trim(cmbrowlist.Text) = "Brand" And Trim(cmbcollist.Text) = "Salesman" Then
    OpenReport CURDIR & "\BRD_MIS.RP1", "Brand Wise Salesman Wise Sales", 2, Trim(SAL_QRY), 6
End If
cancel
End Sub

Private Function CALLIDS()
Dim LVITEM As String
n = 0
ARR = ""
If Trim(cmbcollist.Text) = "Route" Then
    For i = 1 To lv1.ListItems.count
        If lv1.ListItems(i).Checked = True Then
            LVITEM = Trim(lv1.ListItems(i).Text)
            LVITEM = SETMAX(Trim(LVITEM), 15)
            ARR = ARR & LVITEM & ","
            n = n + 1
        End If
    Next
ElseIf Trim(cmbcollist.Text) = "Brand" Or Trim(cmbcollist.Text) = "Segment" Or Trim(cmbcollist.Text) = "Pack Type" Or Trim(cmbcollist.Text) = "Pack Size" Then
    For i = 1 To lv1.ListItems.count
        If lv1.ListItems(i).Checked = True Then
            LVITEM = Trim(lv1.ListItems(i).Text)
            LVITEM = SETMAX(Trim(LVITEM), 15)
            ARR = ARR & LVITEM & ","
            n = n + 1
        End If
    Next
ElseIf Trim(cmbcollist.Text) = "Category" Then
    For i = 1 To lv1.ListItems.count
        If lv1.ListItems(i).Checked = True Then
            LVITEM = Trim(lv1.ListItems(i).Text)
            LVITEM = SETMAX(Trim(LVITEM), 10)
            ARR = ARR & LVITEM & ","
            n = n + 1
        End If
    Next
Else
    For i = 1 To lv1.ListItems.count
        If lv1.ListItems(i).Checked = True Then
            ARR = ARR & Trim(lv1.ListItems(i).ListSubItems(1)) & ","
            n = n + 1
        End If
    Next
End If
If n > 0 Then
    ARR = Mid(Trim(ARR), 1, Len(Trim(ARR)) - 1)
    CALLIDS = ARR
End If
End Function

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.KEY = "sku" Then
    SAL_QRY = "inV->inv_dat >= ctod('" & Format(FDATE.Value, "MM/DD/YYYY") & "').and.inv->inv_dat <= ctod('" & Format(TDATE.Value, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\CMP_RPT.RP1", "SKU Wise Route Sales", 2, Trim(SAL_QRY), 6
End If
End Sub
Private Function SETMAX(ByRef st As String, ML As Integer)
Dim l As Integer
l = Len(Trim(st))
If l > ML Then
    st = Mid(Trim(st), 1, ML)
ElseIf l < ML Then
    For K = l To ML - 1
        st = st & " "
    Next
End If
SETMAX = st
End Function

Private Sub Check8_Click()
Check8_KeyDown 13, 0
End Sub

Private Sub Check8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check8.Value = 1 Then
        cmbCaldates.Enabled = True
        If RSWEK.State = 1 Then RSWEK.Close
        RSWEK.Open "select wek_idy from wek", con, adOpenKeyset, adLockPessimistic
        cmbCaldates.CLEAR
        If Not RSWEK.EOF Then
            While Not RSWEK.EOF
                cmbCaldates.AddItem RSWEK("wek_idy")
                RSWEK.MoveNext
            Wend
            cmbCaldates.SetFocus
        Else
            MsgBox "Sales Calendar does not exist!", vbExclamation, "DAS Version 4.0"
            Check8.SetFocus
            Check8.Value = 0
        End If
    Else
        cmbCaldates.Enabled = False
        FDATE.Enabled = True
        TDATE.Enabled = True
        FDATE.SetFocus
    End If
End If
End Sub

Private Sub cmbCaldates_Click()
    'cmbCaldates_KeyDown 13, 0
End Sub


