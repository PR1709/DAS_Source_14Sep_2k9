VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_BRK1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Breakage Limits"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "M_FRM_BRK1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5145
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Breakage/Leakage Limits"
      TabPicture(0)   =   "M_FRM_BRK1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MS"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtbrk"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton Command1 
         Caption         =   "Load &All SKU's"
         Height          =   375
         Left            =   7005
         TabIndex        =   4
         Top             =   495
         Width           =   1695
      End
      Begin VB.TextBox txtbrk 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   2340
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   4260
         Left            =   45
         TabIndex        =   1
         Top             =   900
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   7514
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   330
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   240
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
            Picture         =   "M_FRM_BRK1.frx":045E
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_BRK1.frx":09A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_BRK1.frx":0D6A
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_BRK1.frx":10BE
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_BRK1.frx":16F6
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_BRK1.frx":1A4A
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_BRK1.frx":1EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_BRK1.frx":219E
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_BRK1.frx":25AA
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_BRK1.frx":26B6
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_BRK1.frx":2A0A
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_BRK1.frx":2E1E
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   8
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ACT_LST"
                  Text            =   "&1 All Accounts List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_CAT"
                  Text            =   "&2 Retailer Category"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_LOC"
                  Text            =   "&3 Retailer Location"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_MAR"
                  Text            =   "&4 Retailer Market"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_ROU"
                  Text            =   "&5 Retailer Beat"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_SMN"
                  Text            =   "&6 Retailer Salesman"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SUP_LST"
                  Text            =   "&7 Supplier List"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RET_ID_LST"
                  Text            =   "&8 Retailer ID List"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
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
            Object.Visible         =   0   'False
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
Attribute VB_Name = "M_FRM_BRK1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsshl As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
Dim rsbrk As New ADODB.Recordset
Dim R As Integer, tmp As Integer
Private Sub Command1_Click()
If rsprd.State = 1 Then rsprd.Close
If rsbrk.State = 1 Then rsbrk.Close
If rsshl.State = 1 Then rsshl.Close
rsprd.Open "select prd_nme,prd_idy,SUB_UNT from prd group by prd_idy", CON, adOpenKeyset, adLockOptimistic
rsshl.Open "select prd_idy,shl_typ,shl_nme from shl", CON, adOpenKeyset, adLockOptimistic
rsbrk.Open "SELECT * FROM BRK", CON, adOpenKeyset, adLockOptimistic
MS.Rows = rsprd.RecordCount + 1
R = 1
While Not rsprd.EOF
    If rsshl.State = 1 Then rsshl.Close
    If rsbrk.RecordCount > 0 Then rsbrk.MoveFirst
    MS.TextMatrix(R, 0) = rsprd("prd_nme")
    MS.TextMatrix(R, 1) = rsprd("SUB_UNT")
    MS.TextMatrix(R, 6) = rsprd("prd_idy")
    rsshl.Open "SELECT SHL_NME FROM SHL WHERE PRD_IDY='" & Trim(rsprd("PRD_IDY")) & "' AND SHL_TYP='Shell'", CON, adOpenKeyset
    'rsshl.Find "prd_idy='" & Trim(rsprd("prd_idy")) & "' and shl_typ='Shell'"
    If rsshl.EOF = False Then MS.TextMatrix(R, 2) = rsshl("shl_nme")
    'If rsshl.RecordCount > 0 Then rsshl.MoveFirst
    If rsshl.State = 1 Then rsshl.Close
    'rsshl.Find "prd_idy='" & Trim(rsprd("prd_idy")) & "' and shl_typ='Pack'"
    rsshl.Open "select shl_nme from shl where prd_idy='" & Trim(rsprd("prd_idy")) & "' and shl_typ='Pack'"
    If rsshl.EOF = False Then MS.TextMatrix(R, 3) = rsshl("shl_nme")
    rsbrk.Find "prd_idy='" & Trim(rsprd("prd_idy")) & "'"
    If rsbrk.EOF = False Then
        MS.TextMatrix(R, 4) = rsbrk("brk_qty")
        MS.TextMatrix(R, 5) = rsbrk("lek_qty")
    End If
    rsprd.MoveNext
    R = R + 1
Wend
If rsprd.RecordCount > 0 Then
    MS.SetFocus
    MS.FixedCols = 4
    MS.Col = 4
    MS.Row = 1
    Command1.Enabled = False
Else
    MsgBox "SKU's not Found!", vbExclamation, "Effmcg"
    Command1.Enabled = True
    Command1.SetFocus
End If
End Sub

Private Sub Form_Activate()
Command1.SetFocus
End Sub

Private Sub loadms()
MS.Cols = 7
MS.TextMatrix(0, 0) = "SKU Name"
MS.TextMatrix(0, 1) = "UPC"
MS.TextMatrix(0, 2) = "Shells"
MS.TextMatrix(0, 3) = "Packs"
MS.TextMatrix(0, 4) = "Breakages"
SCl MS, 4
MS.TextMatrix(0, 5) = "Leakages"
SCl MS, 5
MS.TextMatrix(0, 6) = "Product ID"
MS.ColWidth(0) = 2200
MS.ColWidth(1) = 500
MS.ColWidth(2) = 1500
MS.ColWidth(3) = 1500
MS.ColWidth(4) = 1376
MS.ColWidth(5) = 1280
MS.ColWidth(6) = 0
MS.ColAlignment(0) = 1
MS.ColAlignment(1) = 7
MS.ColAlignment(2) = 1
MS.ColAlignment(3) = 1
MS.ColAlignment(4) = 7
MS.ColAlignment(5) = 7
End Sub

Private Sub save()
If MsgBox("Confirm save?", vbExclamation + vbYesNo, "Effmcg") = vbNo Then
    cancel
    Unload Me
    Exit Sub
End If
If MS.Rows <= 1 Then
    MsgBox "Define SKU Wise Breakage/Leakage Limits!", vbExclamation, "Effmcg"
    Command1.SetFocus
    Exit Sub
End If
tmp = 0
For R = 1 To MS.Rows - 1
    If Val(MS.TextMatrix(R, 4)) > 0 Or Val(MS.TextMatrix(R, 5)) > 0 Then
    tmp = 1
    End If
Next
If tmp = 1 Then
    CON.Execute "DELETE FROM BRK"
    If rsbrk.State = 1 Then rsbrk.Close
    rsbrk.Open "select * from brk", CON, adOpenKeyset, adLockOptimistic
    For R = 1 To MS.Rows - 1
        If Val(MS.TextMatrix(R, 3)) > 0 Or Val(MS.TextMatrix(R, 4)) > 0 Then
            rsbrk.AddNew
            rsbrk("prd_idy") = Trim(MS.TextMatrix(R, 6))
            rsbrk("brk_qty") = Val(MS.TextMatrix(R, 4))
            rsbrk("lek_qty") = Val(MS.TextMatrix(R, 5))
            rsbrk.Update
        End If
    Next
    cancel
    Unload Me
Else
    MsgBox "Define SKU Wise Breakage/Leakage Limits!", vbExclamation, "Effmcg"
    MS.Row = 1
    MS.Col = 4
    MS.SetFocus
End If
End Sub

Private Sub cancel()
MS.Rows = 1
Command1.Enabled = True
Command1.SetFocus
End Sub

Private Sub Form_KeyDown(K As Integer, S As Integer)
If K = 120 Then Shell "c:\windows\calc.exe"
If K > 111 And K < 120 Then If TypeOf M_FRM_BRK1.ActiveControl Is TextBox Then M_FRM_BRK1.ActiveControl.Text = Trim(FUNKEY(K))
    If K = 78 And S = 2 Then
        'newrec
    ElseIf K = 83 And S = 2 Then
        save
    ElseIf K = 81 And S = 2 Then
        cancel
    ElseIf K = 88 And S = 2 Then
        Unload Me
    ElseIf K = 67 And S = 2 Then
        'If Toolbar1.Buttons(12).Enabled = True Then CREATE
    ElseIf K = 68 And S = 2 Then
        delete
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And S = 2 Then
        Openreport CURDIR & "\lists.rp1", "Categories list", 0, "", 6
    ElseIf K = 82 And S = 2 Then
        'MDIForm1.RRReport1.ReportName = App.Path & "\Channels list.rrw"
        'MDIForm1.RRReport1.Destination = 1
        'MDIForm1.RRReport1.Printer = "?"
        'MDIForm1.RRReport1.RunReport 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If M_FRM_BRK1.ActiveControl.name = "Command1" Then
        Unload Me
    Else
        save
    End If
End If
End Sub






























''Dim rspgr As New ADODB.Recordset
''Dim rsbrk As New ADODB.Recordset
''Dim rsshl As New ADODB.Recordset
''Dim R As Integer, MD As Boolean
''Private Sub fdate_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 13 Then
''    tdate.SetFocus
''End If
''End Sub
''
''Private Sub Form_Activate()
''SSTab1.Tab = 0
''SSTab1.TabEnabled(1) = False
'''txnid.SetFocus
''DISPM
''End Sub
''
''Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 78 And Shift = 2 Then
''    newrecord
''ElseIf KeyCode = 83 And Shift = 2 Then
''    saverecord
''ElseIf KeyCode = 81 And Shift = 2 Then
''    cancel
''ElseIf KeyCode = 69 And Shift = 2 Then
''    'If Me.ActiveControl.name = "txnid" Then BRKLIST
''End If
''End Sub
''
''
''
''Private Sub Form_KeyPress(KeyAscii As Integer)
''If KeyAscii = 27 Then
''    'If Me.ActiveControl.name = "txnid" And Len(Trim(txnid)) = 0 Then
''     '   Unload Me
''
''    'Else
''    If SSTab1.Tab = 1 Then
''        If lv.Visible = True And UCase(Trim(Label20)) = UCase("Select Transaction") Then
''            lv.Visible = False
''            SSTab1.Tab = 0
''            txnid.SetFocus
'''        ElseIf lv.Visible = True And UCase(Trim(Label20)) = UCase("Select Salesman") Then
'''            lv.Visible = False
'''            SSTab1.Tab = 0
'''            smnid.SetFocus
''        End If
''    Else
''        saverecord
''    End If
''End If
''End Sub
''
''Private Sub Form_Load()
''SETYEAR Me
''End Sub
''
''Private Sub newrecord()
''cancel
''fdate.SetFocus
''DISPM
''MD = False
''End Sub
''
''Private Sub cancel()
'''txnid.Enabled = True
''clrctr Me
''optCat.Value = 0
''optSeg.Value = 0
''optshl.Value = 0
''fdate.Value = Date
''tdate.Value = Date
''MS.Clear
''MS.Rows = 1
''If SSTab1.Tab = 1 Then
''    lv.Visible = False
''    SSTab1.Tab = 0
''End If
''SSTab1.TabEnabled(1) = False
''DISPM
'''txnid.SetFocus
''End Sub
''
''Private Sub Label2_Click()
''
''End Sub
''
''Private Sub lv_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
''If lv.SortOrder = lvwAscending Then
''lv.SortOrder = lvwDescending
''ElseIf lv.SortOrder = lvwDescending Then
''lv.SortOrder = lvwAscending
''End If
''lv.SortKey = CH.Index - 1
''End Sub
''
''Private Sub lv_DblClick()
''lv_KeyDown 13, 0
''End Sub
''
''Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 13 Then
''    If lv.ListItems.Count > 0 Then
''    If rsbrk.State = 1 Then rsbrk.Close
''        Dim ls As ListItem
''        Set ls = lv.SelectedItem
''        rsbrk.Open "select * from brk where txn_idy like '" & Trim(ls.Text) & "'", CON, adOpenKeyset, adLockOptimistic
''        MD = True
'''        txnid = rsbrk("txn_idy")
''        fdate.Value = rsbrk("frm_dat")
''        tdate.Value = rsbrk("to_dat")
''        MS.Cols = 4
''        MS.ColWidth(0) = 0
''        MS.ColWidth(1) = 1500
''        MS.ColWidth(2) = 1300
''        MS.ColWidth(3) = 1300
''        If Trim(rsbrk("grp_typ")) = "cat" Then
''            optCat.Value = True
''            MS.TextMatrix(0, 0) = "Group ID"
''            MS.TextMatrix(0, 1) = "Category"
''        ElseIf Trim(rsbrk("grp_typ")) = "seg" Then
''            optSeg.Value = True
''            MS.TextMatrix(0, 0) = "Group ID"
''            MS.TextMatrix(0, 1) = "Segment"
''        Else
''            optshl.Value = True
''            MS.ColWidth(0) = 1000
''            MS.TextMatrix(0, 0) = "Shell ID"
''            MS.TextMatrix(0, 1) = "Shell Name"
''        End If
''        MS.TextMatrix(0, 2) = "Breakage Qty"
''        MS.ColAlignment(2) = 7
''        SCl MS, 2
''        MS.TextMatrix(0, 3) = "Leakage Qty"
''        MS.ColAlignment(3) = 7
''        SCl MS, 3
''        R = 1
''        MS.Rows = rsbrk.RecordCount + 1
''        While Not rsbrk.EOF
''            MS.TextMatrix(R, 0) = rsbrk("grp_idy")
''            MS.TextMatrix(R, 1) = rsbrk("grp_nme")
''            MS.TextMatrix(R, 2) = rsbrk("brk_qty")
''            MS.TextMatrix(R, 3) = rsbrk("lek_qty")
''            rsbrk.MoveNext
''            R = R + 1
''        Wend
''        MS.SetFocus
''        MS.Col = 2
''        SSTab1.Tab = 0
''        SSTab1.TabEnabled(1) = False
''        lv.Visible = False
''    End If
''End If
''End Sub
''
''Private Sub MS_KeyPress(KeyAscii As Integer)
''If KeyAscii >= 48 And KeyAscii <= 57 Then
''    If MS.Col = 2 Or MS.Col = 3 Then
''         txtbrk.Visible = True
''         txtbrk.Left = MS.Left + MS.CellLeft
''         txtbrk.Top = MS.Top + MS.CellTop
''         txtbrk.Height = MS.CellHeight
''         txtbrk.Width = MS.CellWidth
''         txtbrk.Visible = True
''         txtbrk.SetFocus
''         txtbrk = Chr(KeyAscii)
''         txtbrk.SelStart = Len(Trim(txtbrk))
''    End If
''End If
''End Sub
''
''Private Sub MS_Scroll()
''If txtbrk.Visible = True Then
''    txtbrk = ""
''    txtbrk.Visible = False
''    MS.SetFocus
''End If
''End Sub
''
''Private Sub optCat_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 13 Then
''    If rspgr.State = 1 Then rspgr.Close
''    rspgr.Open "select grp_idy,grp_cat from pgr GROUP BY grp_cat ORDER BY grp_idy", CON, adOpenKeyset, adLockOptimistic
''    With MS
''        .Rows = rspgr.RecordCount + 1
''        .Cols = 4
''        .TextMatrix(0, 0) = "Group ID"
''        .ColWidth(0) = 0
''        .TextMatrix(0, 1) = "Category"
''        .ColWidth(1) = 1500
''        .ColAlignment(1) = 1
''        .TextMatrix(0, 2) = "Breakage Qty"
''        SCl MS, 2
''        .ColWidth(2) = 1300
''        .ColAlignment(2) = 7
''        .TextMatrix(0, 3) = "Leakage Qty"
''        SCl MS, 3
''        .ColWidth(3) = 1300
''        .ColAlignment(3) = 7
''        R = 1
''        While Not rspgr.EOF
''           .TextMatrix(R, 0) = rspgr("grp_idy")
''           .TextMatrix(R, 1) = Trim(rspgr("grp_cat"))
''           .TextMatrix(R, 2) = 0
''           .TextMatrix(R, 3) = 0
''           R = R + 1
''           rspgr.MoveNext
''        Wend
''        If rspgr.RecordCount > 0 Then
''            MS.SetFocus
''            MS.Col = 2
''            MS.Row = 1
''        End If
''    End With
''End If
''End Sub
''
''Private Sub optSeg_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 13 Then
''    If rspgr.State = 1 Then rspgr.Close
''    rspgr.Open "select grp_idy,grp_cls from pgr group by grp_cls ORDER BY grp_idy", CON, adOpenKeyset, adLockOptimistic
''    If Not rspgr.EOF Then
''        MS.SetFocus
''        MS.Col = 2
''    End If
''    With MS
''        .Rows = rspgr.RecordCount + 1
''        .Cols = 4
''        .TextMatrix(0, 0) = "Group ID"
''        .ColWidth(0) = 0
''        .TextMatrix(0, 1) = "Segment"
''        .ColWidth(1) = 1500
''        .TextMatrix(0, 2) = "Breakage Qty"
''        SCl MS, 2
''        .ColWidth(2) = 1300
''        .ColAlignment(2) = 7
''        .TextMatrix(0, 3) = "Leakage Qty"
''        SCl MS, 3
''        .ColWidth(3) = 1300
''        .ColAlignment(3) = 7
''        R = 1
''        While Not rspgr.EOF
''           .TextMatrix(R, 0) = rspgr("grp_idy")
''           .TextMatrix(R, 1) = rspgr("grp_cls") 'segment
''           .TextMatrix(R, 2) = 0
''           .TextMatrix(R, 3) = 0
''           R = R + 1
''           rspgr.MoveNext
''        Wend
''    End With
''End If
''End Sub
''
''Private Sub optshl_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 13 Then
''    If rsshl.State = 1 Then rsshl.Close
''    rsshl.Open "SELECT DISTINCT(SHL_IDY),SHL_NME FROM SHL", CON, adOpenKeyset, adLockOptimistic
''    If Not rsshl.EOF Then
''        MS.SetFocus
''        MS.Col = 2
''    End If
''    With MS
''        .Rows = rsshl.RecordCount + 1
''        .Cols = 4
''        .TextMatrix(0, 0) = "Shell ID"
''        .ColWidth(0) = 1000
''        .TextMatrix(0, 1) = "Shell Name"
''        .ColWidth(1) = 2000
''        .ColAlignment(1) = 1
''        .TextMatrix(0, 2) = "Breakage Qty"
''        .ColAlignment(2) = 7
''        SCl MS, 2
''        .ColWidth(2) = 1300
''        .TextMatrix(0, 3) = "Leakage Qty"
''        .ColAlignment(3) = 7
''        SCl MS, 3
''        .ColWidth(3) = 1300
''        R = 1
''        While Not rsshl.EOF
''           .TextMatrix(R, 0) = rsshl(0)
''           .TextMatrix(R, 1) = rsshl("shl_nme") 'segment
''           .TextMatrix(R, 2) = 0
''           .TextMatrix(R, 3) = 0
''           R = R + 1
''           rsshl.MoveNext
''        Wend
''    End With
''End If
''End Sub
''
''Private Sub tdate_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 13 Then
''    If tdate.Value < fdate.Value Then
''        MsgBox "Invalid Date", vbExclamation
''        tdate.SetFocus
''    Else
''        optCat.SetFocus
''    End If
''End If
''End Sub
''
''Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
''If Button.KEY = "n" Then
''    newrecord
''ElseIf Button.KEY = "s" Then
''    saverecord
''ElseIf Button.KEY = "c" Then
''    cancel
''ElseIf Button.KEY = "p" Then
''
''ElseIf Button.KEY = "e" Then
''    Unload Me
''ElseIf Button.KEY = "h" Then
''    M_HELP
''ElseIf Button.KEY = "ed" Then
''    'If Me.ActiveControl.name = "txnid" Then
''    '    BRKLIST
''    'End If
''End If
''End Sub
''
'''Private Sub txnid_KeyDown(K As Integer, Shift As Integer)
'''If CHECKTB("brk.dbf") = 1 Then
'''    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
'''    Exit Sub
'''End If
'''Dim rsbrk As New ADODB.Recordset
'''If K = 13 Or K = 40 Then
'''    If rsbrk.State = 1 Then RS.Close
'''    rsbrk.Open "select * from brk where txn_idy like '" & Sincrement(txnid) & "'", CON, adOpenStatic
'''    If rsbrk.RecordCount > 0 Then
'''    MD = True
'''        txnid = rsbrk("txn_idy")
'''        fdate.Value = rsbrk("frm_dat")
'''        tdate.Value = rsbrk("to_dat")
'''        MS.Cols = 4
'''        MS.ColWidth(0) = 0
'''        MS.ColWidth(1) = 1500
'''        MS.ColWidth(2) = 1300
'''        MS.ColWidth(3) = 1300
'''        If Trim(rsbrk("grp_typ")) = "cat" Then
'''            optcat.Value = True
'''            MS.TextMatrix(0, 0) = "Group ID"
'''            MS.TextMatrix(0, 1) = "Category"
'''        ElseIf Trim(rsbrk("grp_typ")) = "seg" Then
'''            optseg.Value = True
'''            MS.TextMatrix(0, 0) = "Group ID"
'''            MS.TextMatrix(0, 1) = "Segment"
'''        Else
'''            optshl.Value = True
'''            MS.ColWidth(0) = 1000
'''            MS.TextMatrix(0, 0) = "Shell ID"
'''            MS.TextMatrix(0, 1) = "Shell Name"
'''        End If
'''        MS.TextMatrix(0, 2) = "Breakage Qty"
'''        MS.ColAlignment(2) = 7
'''        SCl MS, 2
'''        MS.TextMatrix(0, 3) = "Leakage Qty"
'''        MS.ColAlignment(3) = 7
'''        SCl MS, 3
'''        R = 1
'''        MS.Rows = rsbrk.RecordCount + 1
'''        While Not rsbrk.EOF
'''            MS.TextMatrix(R, 0) = rsbrk("grp_idy")
'''            MS.TextMatrix(R, 1) = rsbrk("grp_nme")
'''            MS.TextMatrix(R, 2) = rsbrk("brk_qty")
'''            MS.TextMatrix(R, 3) = rsbrk("lek_qty")
'''            rsbrk.MoveNext
'''            R = R + 1
'''        Wend
'''        MS.SetFocus
'''        MS.Col = 2
'''    Else
'''        BRKLIST
'''    End If
'''    End If
''''md = True
'''End Sub
''
''Private Sub txtbrk_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 13 Then
''    txtbrk.Visible = False
''    MS.Text = Trim(txtbrk)
''    txtbrk = ""
''    If optshl.Value = 1 Then
''        If MS.Col = 4 Then
''            MS.Col = MS.Col + 1
''        Else
''            MS.Row = MS.Row + 1
''            MS.Col = 4
''        End If
''    Else
''        If MS.Col = 2 Then
''            MS.Col = MS.Col + 1
''        Else
''            MS.Row = MS.Row + 1
''            MS.Col = 2
''        End If
''    End If
''    MS.SetFocus
''End If
''End Sub
''Private Sub saverecord()
''If CHECKTB("brk.dbf") = 1 Then
''    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
''    Exit Sub
''End If
''If MsgBox("Save Record ?", vbYesNo + vbExclamation) = vbNo Then
''    cancel
''    Exit Sub
''End If
''If MS.Rows <= 1 Then
''    MsgBox "Select a Group ", vbExclamation
''    optCat.SetFocus
''    Exit Sub
''End If
''If MD = False Then
''    If rsbrk.State = 1 Then rsbrk.Close
''    rsbrk.Open "select * from brk where txn_idy=''", CON, adOpenKeyset, adLockOptimistic
''    'txnid = increment("idy_brk")
''    For R = 1 To MS.Rows - 1
''        If Val(MS.TextMatrix(R, 2)) > 0 And Val(MS.TextMatrix(R, 3)) > 0 Then
''            rsbrk.AddNew
''            'rsbrk("txn_idy") = Trim(txnid)
''            rsbrk("frm_dat") = fdate.Value
''            rsbrk("to_dat") = tdate.Value
''            rsbrk("grp_idy") = MS.TextMatrix(R, 0)
''            rsbrk("grp_nme") = MS.TextMatrix(R, 1)
''            rsbrk("brk_qty") = MS.TextMatrix(R, 2)
''            rsbrk("lek_qty") = MS.TextMatrix(R, 3)
''            If optCat.Value = True Then
''                rsbrk("grp_typ") = "cat"
''            ElseIf optSeg.Value = True Then
''                rsbrk("grp_typ") = "seg"
''            Else
''                rsbrk("grp_typ") = "shl"
''            End If
''            rsbrk.Update
''        End If
''    Next
''    'MsgBox "New Transaction ID is " & Trim(txnid), vbExclamation
''    cancel
''Else
''    If rsbrk.State = 1 Then rsbrk.Close
''    'rsbrk.Open "select * from brk where txn_idy='" & Trim(txnid) & "'", CON, adOpenKeyset, adLockOptimistic
''    For R = 1 To MS.Rows - 1
''        'rsbrk("txn_idy") = Trim(txnid)
''        rsbrk("brk_qty") = MS.TextMatrix(R, 2)
''        rsbrk("lek_qty") = MS.TextMatrix(R, 3)
''        rsbrk.Update
''    Next
''    cancel
''End If
''End Sub
''Private Sub BRKLIST()
''If CHECKTB("brk.dbf") = 1 Then
''    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
''    Exit Sub
''End If
''Dim ls As ListItem
''Dim rsbrk As New ADODB.Recordset
''If rsbrk.State = 1 Then rsbrk.Close
'''rsbrk.Open "select * from brk where txn_idy like '" & Trim(txnid) & "%' group by txn_idy", CON, adOpenStatic
''Label20.Caption = "Select Transaction"
''lv.ColumnHeaders.Clear
''lv.ColumnHeaders.Add , , "Transaction ID", 2000
''lv.ColumnHeaders.Add , , "Group Name", 2500
'''lv.ColumnHeaders.Add , , "Incentive Amt", 1500, 2
''lv.ListItems.Clear
''If rsbrk.RecordCount > 0 Then
''    For A = 1 To rsbrk.RecordCount
''        Set ls = lv.ListItems.Add(, , Trim(rsbrk("txn_idy")))
''        ls.ListSubItems.Add , , Trim(rsbrk("grp_nme"))
''        rsbrk.MoveNext
''    Next
''    SSTab1.TabEnabled(1) = True
''    SSTab1.Tab = 1
''    lv.Visible = True
''    lv.SetFocus
''    MD = True
''Else
''    If SSTab1.Tab = 0 Then
''        MsgBox "Transactions Not Found!", vbExclamation, "Effmcg"
''        SSTab1.Tab = 0
'''        txnid.Enabled = True
'''        txnid.SetFocus
''        SendKeys "{HOME}+{END}"
''        MD = False
''    End If
''End If
''End Sub
''
''Private Sub DISPM()
''MS.Cols = 4
''MS.Rows = 1
''MS.TextMatrix(0, 0) = "Group ID"
''MS.ColWidth(0) = 0
''MS.TextMatrix(0, 1) = "Category"
''MS.ColWidth(1) = 1500
''MS.TextMatrix(0, 2) = "Breakage Qty."
''MS.ColAlignment(2) = 7
''SCl MS, 2
''MS.ColWidth(2) = 1300
''MS.TextMatrix(0, 3) = "Leakage Qty."
''MS.ColAlignment(3) = 7
''SCl MS, 3
''MS.ColWidth(3) = 1300
''End Sub
Private Sub Form_Load()
loadms
End Sub

Private Sub MS_KeyPress(KeyAscii As Integer)
If MS.Col = 4 Or MS.Col = 5 Then
    If (KeyAscii >= 48 And KeyAscii <= 57) Then
        txtbrk = ""
        txtbrk.Visible = True
        txtbrk.Left = MS.Left + MS.CellLeft
        txtbrk.Top = MS.Top + MS.CellTop
        txtbrk.Height = MS.CellHeight
        txtbrk.Width = MS.CellWidth
        txtbrk = Chr(KeyAscii)
        txtbrk.SelStart = Trim(Len(txtbrk))
        txtbrk.SetFocus
    End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
'    newrec
ElseIf Button.KEY = "s" Then
    save
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "d" Then
    delete
ElseIf Button.KEY = "p" Then
'    Openreport CURDIR & "\lists.rp1", "Categories list", 0, "", 6
ElseIf Button.KEY = "pr" Then
    'MDIForm1.RRReport1.WindowTitle = "Retailers' Categories List Report"
    'MDIForm1.RRReport1.ReportName = App.Path & "\Channels list.rrw"
    'MDIForm1.RRReport1.Destination = 1
    'MDIForm1.RRReport1.Printer = "?"
    'MDIForm1.RRReport1.RunReport 1
ElseIf Button.KEY = "crt" Then
'    CREATE
End If

End Sub

Private Sub txtbrk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Val(txtbrk) > Val(MS.TextMatrix(MS.Row, 1)) Then
        MsgBox "Sub Units are " & Val(MS.TextMatrix(MS.Row, 1))
        txtbrk.SetFocus
        SendKeys "{home}+{end}"
    Else
        MS.Text = Trim(txtbrk)
        txtbrk = ""
        txtbrk.Visible = False
        If MS.Col = 4 Then
            MS.Col = MS.Col + 1
        Else
            If MS.Row <> MS.Rows - 1 Then
                MS.Row = MS.Row + 1
                MS.Col = 4
            End If
        End If
    End If
End If
End Sub
Private Sub delete()
If MsgBox("Confirm Delete ?", vbExclamation + vbYesNo, "Effmcg") = vbYes Then
    CON.Execute "delete from brk"
End If
End Sub

Private Sub txtbrk_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
Else
    KeyAscii = 0
End If
End Sub
