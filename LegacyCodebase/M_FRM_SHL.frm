VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_SHL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Returnables - Shell Identification (Shell Description)"
   ClientHeight    =   5010
   ClientLeft      =   1815
   ClientTop       =   1650
   ClientWidth     =   8880
   Icon            =   "M_FRM_SHL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8017
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Shell - Crate Details"
      TabPicture(0)   =   "M_FRM_SHL.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LV2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_SHL.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label31"
      Tab(1).Control(1)=   "LV1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3735
         Begin VB.TextBox shl_idy 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   1185
         End
         Begin VB.TextBox unt_rat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   330
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1185
         End
         Begin VB.TextBox shl_nme 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   25
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   720
            Width           =   2385
         End
         Begin VB.ListBox SHL_TYP 
            Appearance      =   0  'Flat
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
            Height          =   270
            ItemData        =   "M_FRM_SHL.frx":047A
            Left            =   1200
            List            =   "M_FRM_SHL.frx":0481
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   15
            Top             =   315
            Width           =   165
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   675
            TabIndex        =   14
            Top             =   1245
            Width           =   360
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   615
            TabIndex        =   13
            Top             =   795
            Width           =   420
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Rate"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   1755
            Width           =   675
         End
      End
      Begin VB.CommandButton command1 
         Caption         =   "&Select All"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   4060
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   3255
         Left            =   -74475
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5741
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
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
      Begin MSComctlLib.ListView LV2 
         Height          =   3255
         Left            =   3960
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5741
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Group n Identify SKUs for this Shell - Crate"
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
         Left            =   3960
         TabIndex        =   6
         Top             =   405
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Shell - Crate"
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
         Height          =   330
         Left            =   -74475
         TabIndex        =   1
         Top             =   645
         Width           =   7935
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   1200
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
            Picture         =   "M_FRM_SHL.frx":048C
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SHL.frx":09D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SHL.frx":0D98
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SHL.frx":10EC
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SHL.frx":1724
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SHL.frx":1A78
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SHL.frx":1ED8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SHL.frx":21CC
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SHL.frx":25D8
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SHL.frx":26E4
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SHL.frx":2A38
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_SHL.frx":2E4C
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
      Width           =   8880
      _ExtentX        =   15663
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
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "a"
                  Text            =   "Shells List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Key             =   "b"
                  Text            =   "Packs List"
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
            Object.Visible         =   0   'False
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
Attribute VB_Name = "M_FRM_SHL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim f As Field
Dim RS1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim id  As Integer, id1 As Integer, CH As Integer, sd As Integer
Dim STYPE As String, crd As String
Dim rspgr As New ADODB.Recordset

Private Sub Command1_Click()
If Command1.Caption = "&Select All" Then
    For sd = 1 To LV2.ListItems.count
        LV2.ListItems(sd).Checked = True
    Next
    Command1.Caption = "&DeSelect All"
Else
    For sd = 1 To LV2.ListItems.count
        LV2.ListItems(sd).Checked = False
    Next
    Command1.Caption = "&Select All"
End If
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
If shl_idy.Enabled = True Then
    shl_idy.SetFocus
End If
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 78 And s = 2 Then
    newrec
ElseIf K = 83 And s = 2 Then
    save
ElseIf K = 81 And s = 2 Then
    cancel
ElseIf K = 80 And s = 2 Then
    'PRINT
ElseIf K = 112 Then
    M_HELP
ElseIf K = 68 And s = 2 Then
    DEL_REC
ElseIf K = 88 And s = 2 Then
    Unload Me
ElseIf K = 69 And s = 2 Then
    DISPLIST
ElseIf K = 27 Then
    If M_FRM_SHL.ActiveControl.name = "LV2" Then
        shl_nme.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        SSTab1.Tab = 0
        If shl_idy.Enabled = True Then
            shl_idy.SetFocus
        Else
            shl_nme.SetFocus
        End If
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If Me.ActiveControl.name = "shl_idy" Then
        Unload Me
    Else
        save
    End If
End If
End Sub

Private Sub Form_Load()
checkdongle
If CHECKTB("shl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If rs.State = 1 Then rs.Close
rs.Open "select * from shl", con, adOpenKeyset, adLockPessimistic
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
LV1_KeyDown 13, 1
End Sub

Private Sub LV1_KeyDown(K As Integer, s As Integer)
Dim LS1 As ListItem
If K = 13 Then
    If CHECKTB("shl.dbf,prd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If

    If lv1.ListItems.count > 0 Then
       If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * from shl where shl_IDY like'" & Trim(lv1.SelectedItem) & "'", con, adOpenDynamic, adLockPessimistic
            shl_idy = RS1("shl_idy")
            shl_nme = Trim(RS1("shl_nme"))
            If Trim(RS1!SHL_TYP) = "Pack" Then
                SHL_TYP.Selected(1) = True
            Else
                SHL_TYP.Selected(0) = True
                SHL_TYP = Trim(RS1!SHL_TYP)
            End If
        STYPE = RS1("SHL_TYP")
        prd_idy = RS1("prd_idy")
        unt_rat = convert(RS1("unt_rat"))
        shl_idy.Enabled = False
        lv1.Visible = False
        SSTab1.Tab = 0
        st = ""
        If rs2.State = 1 Then rs2.Close
            rs2.Open "SELECT * FROM SHL WHERE SHL_IDY LIKE '" & Trim(shl_idy.Text) & "'", con, adOpenKeyset, adLockPessimistic
        For i = 0 To rs2.RecordCount - 1
            st = st & "'" & rs2("PRD_IDY") & "'" & ","
            rs2.MoveNext
        Next
        If Trim(st) <> "" Then
            st = Mid(Trim(st), 1, Len(Trim(st)) - 1)
        End If
        LV2.View = lvwReport
        Label3.Visible = True
        LV2.Visible = True
        Command1.Visible = True
        LV2.ColumnHeaders.CLEAR
        LV2.ListItems.CLEAR
        LV2.Refresh
        With LV2.ColumnHeaders
            .Add , , "Product Name", 3100
            .Add , , "Category", 2000
            .Add , , "Segment", 2000
            .Add , , "Product ID"
            .Add , , "Sub Units", 1500, 1
            .Add , , "Short Name", 2500
        End With
        If rs2.State = 1 Then rs2.Close
        If Trim(st) <> "" Then
        rs2.Open "SELECT * FROM PRD WHERE PRD_IDY IN(SELECT PRD_IDY FROM SHL WHERE SHL_IDY LIKE '" & Trim(shl_idy.Text) & "') AND (PRD.PRD_PCK <> '' AND PRD.PRD_PKT <> '')", con, adOpenKeyset, adLockPessimistic
            crd = st
            id1 = 1
'            For I = 0 To rs2.RecordCount - 1
'                Set LS1 = LV2.ListItems.Add(, , rs2("prd_nme"))
'                If rspgr.State = 1 Then rspgr.Close
'                rspgr.Open "select * from pgr where grp_idy='" & rs2("grp_idy") & "'", CON, adOpenKeyset, adlockpessimistic
'                If Not rspgr.EOF Then
'                    LS1.ListSubItems.Add , , rspgr("grp_cat")
'                    LS1.ListSubItems.Add , , rspgr("grp_cls")
'                Else
'                    LS1.ListSubItems.Add , , " "
'                    LS1.ListSubItems.Add , , " "
'                End If
'                LS1.ListSubItems.Add , , rs2("prd_idy")
'                LS1.ListSubItems.Add , , rs2("sub_unt")
'                LS1.ListSubItems.Add , , rs2("prd_prn")
'                rs2.MoveNext
'                LV2.ListItems(I + 1).Checked = True
'            Next
        shl_nme.SetFocus
        End If
End If
    st = ""
    If RS1.State = 1 Then RS1.Close
    RS1.Open "SELECT * FROM SHL WHERE SHL_TYP LIKE '" & Trim(STYPE) & "'", con, adOpenKeyset, adLockPessimistic
    For i = 0 To RS1.RecordCount - 1
        st = st & "'" & RS1("PRD_IDY") & "'" & ","
        RS1.MoveNext
    Next
    If Trim(st) <> "" Then
        st = Mid(Trim(st), 1, Len(Trim(st)) - 1)
    End If
    If rs2.State = 1 Then rs2.Close
        If Trim(st) <> "" Then
        rs2.Open "SELECT * FROM PRD WHERE PRD_IDY NOT IN(SELECT PRD_IDY FROM SHL WHERE SHL_TYP LIKE '" & Trim(STYPE) & "') AND (PRD.PRD_PCK <> '' AND PRD.PRD_PKT <> '')", con, adOpenKeyset, adLockPessimistic
        For i = 0 To rs2.RecordCount - 1
        Set LS1 = LV2.ListItems.Add(, , rs2("prd_nme"))
        If rspgr.State = 1 Then rspgr.Close
        rspgr.Open "SELECT * FROM PGR WHERE GRP_IDY='" & rs2("GRP_IDY") & "'", con, adOpenKeyset, adLockPessimistic
        If Not rspgr.EOF Then
            LS1.ListSubItems.Add , , rspgr("GRP_CAT")
            LS1.ListSubItems.Add , , rspgr("GRP_CLS")
        Else
            LS1.ListSubItems.Add , , " "
            LS1.ListSubItems.Add , , " "
        End If
        LS1.ListSubItems.Add , , rs2("PRD_IDY")
        LS1.ListSubItems.Add , , rs2("SUB_UNT")
        LS1.ListSubItems.Add , , rs2("PRD_PRN")
        rs2.MoveNext
    Next
    End If
End If
id = 0
End Sub

Private Sub prd_idy_KeyDown(K As Integer, s As Integer)
Dim ls As ListItem
If K = 13 Then
End If
If K = 38 Then
shl_nme.SetFocus
End If
End Sub


Private Sub LV2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If LV2.SortOrder = lvwAscending Then
    LV2.SortOrder = lvwDescending
ElseIf LV2.SortOrder = lvwDescending Then
    LV2.SortOrder = lvwAscending
End If
LV2.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub LV2_KeyDown(K As Integer, s As Integer)
CH = 1
If K = 13 Then
    unt_rat.SetFocus
End If
End Sub

Private Sub shl_idy_KeyDown(K As Integer, s As Integer)
CH = 0
Dim ls As ListItem
If K = 13 Then
    DISPLIST
End If
End Sub

Private Sub shl_idy_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub shl_nme_Change()
CH = 1
End Sub

Private Sub shl_nme_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub shl_nme_KeyDown(K As Integer, s As Integer)
If K = 13 Or K = 40 Then
    If Len(Trim(shl_nme)) <> 0 Then
        If SHL_TYP.Enabled = True Then
            SHL_TYP.SetFocus
            SHL_TYP.ListIndex = 0
        Else
            If LV2.Visible = True And LV2.Enabled = True Then LV2.SetFocus
        End If
    Else
        MsgBox "Name Should not be blank ,Define Name!", vbExclamation, "DAS Version 4.0"
        shl_nme.SetFocus
    End If
ElseIf K = 38 Then
    If shl_idy.Enabled = True Then
        If Len(Trim(shl_nme)) <> 0 Then
            shl_idy.SetFocus
        Else
            MsgBox "Name Should not be blank ,Define Name!", vbExclamation, "DAS Version 4.0"
            shl_nme.SetFocus
        End If
    End If
End If
End Sub

Private Sub shl_nme_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub SHL_TYP_KeyDown(K As Integer, s As Integer)
CH = 1
If K = 13 Then
    If CHECKTB("shl.dbf,prd.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        cancel
        Exit Sub
    End If

    'If id = 1 Then
        Label3.Visible = True
        LV2.Visible = True
        Command1.Visible = True
        st = ""
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM SHL WHERE SHL_TYP LIKE '" & Trim(SHL_TYP.Text) & "'", con, adOpenKeyset, adLockPessimistic
        For i = 0 To RS1.RecordCount - 1
            st = st & "'" & RS1("PRD_IDY") & "'" & ","
            RS1.MoveNext
        Next
        If Trim(st) <> "" Then
            st = Mid(Trim(st), 1, Len(Trim(st)) - 1)
        End If
        LV2.View = lvwReport
        LV2.ListItems.CLEAR
        Label3.Visible = True
        LV2.Visible = True
        Command1.Visible = True
        If rs2.State = 1 Then rs2.Close
            rs2.Open "SELECT * FROM PRD WHERE PRD_IDY NOT IN(SELECT PRD_IDY FROM SHL WHERE SHL_TYP LIKE '" & Trim(SHL_TYP.Text) & "') AND (PRD.PRD_PCK <> '' AND PRD.PRD_PKT <> '')", con, adOpenKeyset, adLockPessimistic
            With LV2.ColumnHeaders
            .Add , , "Product Name", 3100
            .Add , , "Category", 2000
            .Add , , "Segment", 2000
            .Add , , "Product ID"
            .Add , , "Sub Units", 1500, 1
            .Add , , "Short Name", 2500
        End With
        If rs2.RecordCount > 0 Then
            For i = 0 To rs2.RecordCount - 1
                Set ls = LV2.ListItems.Add(, , rs2("prd_nme"))
                If rspgr.State = 1 Then rspgr.Close
                rspgr.Open "select * from pgr where grp_idy='" & rs2("grp_idy") & "'", con, adOpenKeyset, adLockPessimistic
                If Not rspgr.EOF Then
                    ls.ListSubItems.Add , , rspgr("grp_cat")
                    ls.ListSubItems.Add , , rspgr("grp_cls")
                Else
                    ls.ListSubItems.Add , , " "
                    ls.ListSubItems.Add , , " "
                End If
                ls.ListSubItems.Add , , rs2("prd_idy")
                ls.ListSubItems.Add , , rs2("sub_unt")
                ls.ListSubItems.Add , , rs2("prd_prn")
                rs2.MoveNext
            Next
        End If
    'End If
    LV2.Enabled = True
    If LV2.ListItems.count > 0 Then LV2.SetFocus
ElseIf K = 38 And s = 1 Then
    shl_nme.SetFocus
ElseIf K = 40 And s = 1 Then
    unt_rat.SetFocus
End If
End Sub

Private Sub SSTab1_KeyDown(K As Integer, s As Integer)
If K = 27 Then Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    newrec
ElseIf Button.KEY = "s" Then
    save
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "e" Then
    exitfrm
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If Me.ActiveControl.name = "shl_idy" Then
        DISPLIST
    End If
ElseIf Button.KEY = "d" Then
    DEL_REC
End If
End Sub
Public Sub newrec()
shl_idy = ""
shl_nme = ""
prd_idy = ""
unt_rat = ""
shl_idy.Enabled = False
id = 1
id1 = 0
SHL_TYP.Enabled = True
SHL_TYP.ListIndex = 0
shl_idy = ""
shl_nme = ""
prd_idy = ""
unt_rat = ""
shl_idy.Enabled = True
LV2.ListItems.CLEAR
Label3.Visible = False
LV2.Visible = False
LV2.ColumnHeaders.CLEAR
LV2.ListItems.CLEAR
Command1.Visible = False
SSTab1.Tab = 0
shl_nme.SetFocus
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.KEY = "a" Then
    OpenReport CURDIR & "\lists.RP1", "Shells List", 2, "SHL_TYP='Shell'", 6
Else
    OpenReport CURDIR & "\lists.RP1", "Packs Lists", 2, "SHL_TYP='Pack'", 6
End If
End Sub

Private Sub unt_rat_GotFocus()
SendKeys "{HOME}+{END}"
End Sub

Private Sub unt_rat_KeyDown(K As Integer, s As Integer)
CH = 1
If K = 13 Then
    save
ElseIf K = 38 Then
    If SHL_TYP.Enabled = True Then
        SHL_TYP.SetFocus
    Else
        shl_nme.SetFocus
    End If
End If
End Sub
Public Sub cancel()
shl_idy = ""
shl_nme = ""
prd_idy = ""
unt_rat = ""
shl_idy.Enabled = True
shl_idy.SetFocus
LV2.ListItems.CLEAR
Label3.Visible = False
LV2.Visible = False
LV2.ColumnHeaders.CLEAR
LV2.ListItems.CLEAR
Command1.Visible = False
SSTab1.Tab = 0
End Sub
Public Sub DEL_REC()
If CHECKTB("shl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

If MsgBox("Are you sure you want to delete Shell ?", vbYesNo + vbCritical, "DAS Version 4.0") = vbYes Then
    If check("SHL", "SHL_IDY", shl_idy) = False Then
        MsgBox "Shell Not Found, Define Shell !", vbExclamation, "DAS Version 4.0"
        Exit Sub
    Else
    ' Check for TWP Version NO SAVE
    Dim DIR_RS As New ADODB.Recordset
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'MDL' AND DIR_IDX = 'TWP'"), con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        If Trim(DIR_RS("DIR_ALS")) = "Y" Then
            MsgBox ("Cannot Delete Master Data for KO-TWP Version!"), vbCritical, "DAS Version 4.0"
            cancel
            Exit Sub
        Else
        If deletecheck = False Then
            con.Execute "DELETE FROM SHL WHERE SHL_IDY LIKE '" & Trim(shl_idy) & "'"
            Label3.Visible = False
            LV2.Visible = False
            Command1.Visible = False
            cancel
        Else
            MsgBox "Cannot Delete, Transactions Entries Found!", vbExclamation, "DAS Version 4.0"
            If shl_idy.Enabled = True Then
            shl_idy.SetFocus
            Else
                shl_nme.SetFocus
            End If
            Exit Sub
        End If
        End If
    End If
    End If
Else
    cancel
End If
End Sub
Private Function deletecheck() As Boolean
If CHECKTB("snv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Function
End If

If check("SNV", "SHL_idy", Trim(shl_idy)) = False Then
    deletecheck = False
Else
    deletecheck = True
End If
End Function

Public Sub save()
If MsgBox("Confirm Record Save ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
' Check for TWP Version NO SAVE
Dim DIR_RS As New ADODB.Recordset
If DIR_RS.State = 1 Then DIR_RS.Close
DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'MDL' AND DIR_IDX = 'TWP'"), con, adOpenKeyset, adLockPessimistic
If DIR_RS.RecordCount > 0 Then
    If Trim(DIR_RS("DIR_ALS")) = "Y" Then
        MsgBox ("Cannot Save Master Data for KO-TWP Version!"), vbCritical, "DAS Version 4.0"
        cancel
        Exit Sub
    End If
End If
'

If CHECKTB("shl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If id = 1 Then
    If Trim(shl_nme.Text) = "" Then
        MsgBox "Shell Name Should Not Be Blank", vbOKOnly + vbInformation, "Shells/Packs"
        shl_nme.SetFocus
        Exit Sub
    ElseIf id1 = 1 Then
        'CON.Execute "DELETE FROM SHL WHERE PRD_IDY IN(" & Trim(crd) & ") AND SHL_TYP LIKE '" & Trim(STYPE) & "'"
    End If
    LD = increment("IDY_SHL")
    chk_flg = 0
    For i = 1 To LV2.ListItems.count
        If LV2.ListItems(i).Checked = True Then
            rs.AddNew
            rs("shl_idy") = Trim(LD)
            rs("shl_nme") = Trim(shl_nme)
            rs("prd_idy") = Trim(LV2.ListItems(i).ListSubItems(3))
            rs("unt_rat") = Val(unt_rat)
            rs("shl_typ") = Trim(SHL_TYP.Text)
            rs.Update
            chk_flg = 1
        End If
    Next
'    If chk_flg = 0 Then
'        RS.AddNew
'        RS("shl_idy") = Trim(LD)
'        RS("shl_nme") = Trim(shl_nme)
'        RS("prd_idy") = ""
'        RS("unt_rat") = Val(unt_rat)
'        RS("shl_typ") = Trim(SHL_TYP.Text)
'        RS.Update
'    End If
    MsgBox "New " & SHL_TYP & " ID is  " & LD, vbOKOnly + vbInformation, "Record Saved"
    Label3.Visible = False
    LV2.Visible = False
    Command1.Visible = False
    LV2.ListItems.CLEAR
    LV2.Refresh
    shl_idy.Enabled = True
    cancel
Else
    If Trim(shl_nme.Text) = "" Then
        MsgBox "Shell Name Should Not Be Blank", vbOKOnly + vbInformation, "Shells/Packs"
        shl_nme.SetFocus
    Else
        chk_flg = 0
        For i = 1 To LV2.ListItems.count
            If LV2.ListItems(i).Checked = True Then
                rs.AddNew
                rs("shl_idy") = Trim(shl_idy.Text)
                rs("shl_nme") = Trim(shl_nme)
                rs("prd_idy") = Trim(LV2.ListItems(i).ListSubItems(3))
                rs("unt_rat") = Val(unt_rat)
                rs("shl_typ") = Trim(SHL_TYP.Text)
                rs.Update
                chk_flg = 1
            End If
        Next
        Label3.Visible = False
        LV2.Visible = False
        Command1.Visible = False
        LV2.ListItems.CLEAR
        LV2.Refresh
        shl_idy.Enabled = True
        cancel
    End If
End If
id = 0
End Sub
Public Sub DISPLIST()
If CHECKTB("shl.dbf,prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
lv1.ListItems.CLEAR
lv1.ColumnHeaders.CLEAR
lv1.View = lvwReport
With lv1
    .ColumnHeaders.Add 1, , "Shell ID", 1600
    .ColumnHeaders.Add 2, , "Shell Name", 3600
    .ColumnHeaders.Add 3, , "Shell Type", 1600
    .ColumnHeaders.Add 4, , "Unit Rate", 1600, 1
End With
If RS1.State = 1 Then RS1.Close
    If Trim(shl_idy.Text) = "" Then
        RS1.Open "SELECT DISTINCT(SHL_IDY),SHL_NME,SHL_TYP,UNT_RAT FROM SHL", con, adOpenKeyset, adLockPessimistic
        If RS1.RecordCount > 0 Then
        SHL_TYP.Enabled = False
        lv1.Visible = True
        Label3.Visible = False
        LV2.Visible = False
        Command1.Visible = False
        lv1.SetFocus
        SSTab1.Tab = 1
        lv1.ListItems.CLEAR
        RS1.MoveFirst
            For i = 0 To RS1.RecordCount - 1
                Set ls = lv1.ListItems.Add(, , RS1.Fields(0))
                ls.ListSubItems.Add , , RS1.Fields("shl_nme")
                ls.ListSubItems.Add , , RS1.Fields("shl_typ")
                ls.ListSubItems.Add , , RS1.Fields("unt_rat")
            RS1.MoveNext
            Next
        Else
            MsgBox "Shell ID Not Found ! Define Shell ", vbExclamation, "Shells 'n' Packs"
            shl_idy.SetFocus
        End If
    ElseIf Not IsNumeric(shl_idy.Text) Then
        RS1.Open "SELECT DISTINCT(SHL_IDY),SHL_NME,SHL_TYP,UNT_RAT FROM SHL where shl_nme like '" & Trim(shl_idy.Text) & "%'", con, adOpenKeyset, adLockPessimistic
        If RS1.RecordCount > 0 Then
            SHL_TYP.Enabled = False
            lv1.Visible = True
            Label3.Visible = False
            LV2.Visible = False
            Command1.Visible = False
            lv1.SetFocus
            SSTab1.Tab = 1
            lv1.ListItems.CLEAR
            RS1.MoveFirst
            For i = 0 To RS1.RecordCount - 1
                Set ls = lv1.ListItems.Add(, , RS1.Fields(0))
                ls.ListSubItems.Add , , RS1.Fields("shl_nme")
                ls.ListSubItems.Add , , RS1.Fields("shl_typ")
                ls.ListSubItems.Add , , RS1.Fields("unt_rat")
                RS1.MoveNext
            Next
        Else
            MsgBox "Shell ID Not Found ! Define Shell ", vbExclamation, "Shells 'n' Packs"
            shl_idy.SetFocus
            SHL_TYP.ForeColor = vbBlack
        End If
    Else
    If RS1.State = 1 Then RS1.Close
    RS1.Open "SELECT * FROM SHL WHERE SHL_IDY LIKE '" & Sincrement(shl_idy) & "'", con, adOpenKeyset, adLockPessimistic
    ST1 = ""
    If RS1.RecordCount > 0 Then
        SHL_TYP.Enabled = False
        shl_idy.Enabled = False
        STYPE = RS1("SHL_TYP")
        SHL_TYP.ForeColor = vbWhite
        shl_nme = Trim(RS1("SHL_NME"))
        If Trim(RS1!SHL_TYP) = "Pack" Then
            SHL_TYP.Selected(1) = True
            SHL_TYP.ListIndex = 1
            SHL_TYP.Text = SHL_TYP.List(1)
        Else
            SHL_TYP.Selected(0) = True
            SHL_TYP.ListIndex = 0
            SHL_TYP.Text = SHL_TYP.List(i)
        End If
        unt_rat = convert(RS1("UNT_RAT") & "")
        shl_idy = RS1("SHL_IDY")
    End If
    If RS1.State = 1 Then RS1.Close
        RS1.Open "select * from prd where prd_idy in(SELECT PRD_IDY FROM SHL WHERE SHL_IDY LIKE '" & Sincrement(shl_idy) & "')", con, adOpenKeyset, adLockPessimistic
        If RS1.RecordCount > 0 Then
            SHL_TYP.Enabled = False
            lv1.Visible = False
            Label3.Visible = True
            LV2.Visible = True
            Command1.Visible = True
            SSTab1.Tab = 0
            LV2.View = lvwReport
            LV2.ColumnHeaders.CLEAR
            With LV2.ColumnHeaders
                .Add , , "Product Name", 3100
                'CHANGES
                .Add , , "Category", 2000
                .Add , , "Segment", 2000
                .Add , , "Product ID"
                .Add , , "Sub Units", 1500, 1
                .Add , , "Short Name", 2500
            End With
            For i = 0 To RS1.RecordCount - 1
                Set ls = LV2.ListItems.Add(, , RS1.Fields("prd_nme"))
                If rspgr.State = 1 Then rspgr.Close
                rspgr.Open "select * from pgr where grp_idy='" & RS1("grp_idy") & "'", con, adOpenKeyset, adLockPessimistic
                If Not rspgr.EOF Then
                    ls.ListSubItems.Add , , rspgr("grp_cat")
                    ls.ListSubItems.Add , , rspgr("grp_cls")
                Else
                    ls.ListSubItems.Add , , " "
                    ls.ListSubItems.Add , , " "
                End If
                ls.ListSubItems.Add , , RS1.Fields("prd_idy")
                ls.ListSubItems.Add , , RS1.Fields("sub_unt")
                ls.ListSubItems.Add , , RS1.Fields("prd_prn")
                RS1.MoveNext
                LV2.ListItems(i + 1).Checked = True
            Next
            If RS1.State = 1 Then RS1.Close
            RS1.Open "SELECT * FROM SHL WHERE SHL_TYP LIKE '" & Trim(STYPE) & "'", con, adOpenKeyset, adLockPessimistic
            For i = 0 To RS1.RecordCount - 1
                ST1 = ST1 & "'" & RS1("PRD_IDY") & "'" & ","
                RS1.MoveNext
            Next
            If Trim(ST1) <> "" Then
                ST1 = Mid(Trim(ST1), 1, Len(Trim(ST1)) - 1)
            End If
    If rs2.State = 1 Then rs2.Close
            rs2.Open "SELECT * FROM PRD WHERE PRD_IDY NOT IN(SELECT PRD_IDY FROM SHL WHERE SHL_TYP LIKE '" & Trim(STYPE) & "')", con, adOpenKeyset, adLockPessimistic
            For i = 0 To rs2.RecordCount - 1
                Set LS1 = LV2.ListItems.Add(, , rs2("prd_nme"))
                    LS1.ListSubItems.Add , , rs2("prd_idy")
                    LS1.ListSubItems.Add , , rs2("sub_unt")
                    LS1.ListSubItems.Add , , rs2("prd_prn")
                rs2.MoveNext
            Next
    shl_nme.SetFocus
    Else
        MsgBox "Shell SKU's Not Found", vbOKOnly + vbExclamation, "Shells/Packs"
        SHL_TYP.ForeColor = vbBlack
        If shl_idy.Enabled = True Then
            shl_idy.SetFocus
        Else
            shl_nme.SetFocus
        End If
        SendKeys "+{home}", True
    End If
End If
End Sub
Public Sub exitfrm()
Unload Me
End Sub

Private Sub unt_rat_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
    K = 0
End If
End Sub
