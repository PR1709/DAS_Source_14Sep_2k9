VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_PGR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Categories"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "M_FRM_PGR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
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
            Picture         =   "M_FRM_PGR.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PGR.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PGR.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PGR.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PGR.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PGR.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PGR.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PGR.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PGR.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PGR.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PGR.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_PGR.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5200
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   430
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   9155
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "&Category"
      TabPicture(0)   =   "M_FRM_PGR.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_PGR.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "list1"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74640
         TabIndex        =   10
         Top             =   4560
         Width           =   7935
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   150
            Width           =   1815
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Segment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   150
            Value           =   1  'Checked
            Width           =   2415
         End
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1905
         MaxLength       =   25
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2745
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1905
         MaxLength       =   25
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2745
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   720
         Width           =   1065
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   945
         Visible         =   0   'False
         Width           =   7905
         _ExtentX        =   13944
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product Category"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product Segment"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Product Category"
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
         Left            =   -74640
         TabIndex        =   8
         Top             =   600
         Width           =   7905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   675
         TabIndex        =   6
         Top             =   795
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   285
         TabIndex        =   5
         Top             =   1275
         Width           =   1230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Segment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   255
         TabIndex        =   4
         Top             =   1755
         Width           =   1230
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8850
      _ExtentX        =   15610
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
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Category"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "M_FRM_PGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD1 As Boolean, idnew As Boolean
Dim ls As ListItem

Private Sub CREATE()
MOD1 = False
Text1 = ""
save
End Sub

Public Sub delete()
If CHECKTB("pgr.dbf,prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If MsgBox("Delete Product Group ?", vbYesNo + vbCritical, "DAS Version 4.0") = vbYes Then
    If check("pgr", "grp_idy", Text1) = False Then
        MsgBox "Product Group Not Found, Define Product Group !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        clrctr M_FRM_PGR
        Text1.Enabled = True
        Text1.SetFocus
        Exit Sub
    Else
        If check("prd", "grp_idy", Trim(Text1)) = False Then
            ' Check for TWP Version NO DELETE
            Dim DIR_RS As New ADODB.Recordset
            If DIR_RS.State = 1 Then DIR_RS.Close
            DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'MDL' AND DIR_IDX = 'TWP'"), con, adOpenKeyset, adLockPessimistic
            If DIR_RS.RecordCount > 0 Then
                If Trim(DIR_RS("DIR_ALS")) = "Y" Then
                    MsgBox ("Cannot Delete Master Data for KO-TWP Version!"), vbCritical, "DAS Version 4.0"
                    cancel
                    Exit Sub
                Else
                    con.Execute "delete from pgr where grp_idy like '" & Sincrement(Text1) & "'"
                    SSTab1.Tab = 0
                    Text1.Enabled = True
                    Text1.SetFocus
                    clrctr M_FRM_PGR
                End If
            End If
        Else
            MsgBox "Cannot Delete, Transaction Exists", vbExclamation, "DAS Version 4.0"
            Text2.SetFocus
            Exit Sub
        End If
    End If
    cancel
Else
    cancel
End If
End Sub

Private Sub prdgrplist()
If CHECKTB("pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If Check1.Value = 1 Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from pgr where grp_cat like '" & Text1 & "%'", con, adOpenKeyset, adLockPessimistic
Else
    If rs.State = 1 Then rs.Close
    rs.Open "select * from pgr where grp_cls like '" & Text1 & "%'", con, adOpenKeyset, adLockPessimistic
End If
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("GRP_CAT")))
        ls.ListSubItems.Add , , Trim(rs("GRP_CLS"))
        ls.ListSubItems.Add , , Trim(rs("GRP_IDY"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.ListItems(1).Selected = True 'Added by PNR
    list1.SetFocus
Else
    MsgBox "Category Not Found, Define Category !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub save()
Dim rs As New ADODB.Recordset

If MsgBox("Confirm Record Save ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbNo Then
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
If CHECKTB("pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If

If Len(Trim(Text2)) = 0 Then
MsgBox "Category Should Not be Blank, Enter Category !", vbExclamation, "DAS Version 4.0"
Text1.Enabled = False
SSTab1.Tab = 0
Text2.SetFocus
Exit Sub
End If

If Len(Trim(Text3)) = 0 Then
MsgBox "Segment Should Not be Blank, Enter Segment !", vbExclamation, "DAS Version 4.0"
Text1.Enabled = False
SSTab1.Tab = 0
Text3.SetFocus
Exit Sub
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from pgr where grp_cat like '" & Trim(Text2) & "' and grp_cls like '" & Trim(Text3) & "'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If rs("grp_idy") <> Trim(Text1) Then
        MsgBox "Duplicate Values are Invalid!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text3.SetFocus
        Exit Sub
    End If
End If
If rs.State = 1 Then rs.Close
If MOD1 = True Then
    rs.Open "select * from pgr where grp_idy like '" & Trim(Text1) & "'", con, adOpenKeyset, adLockPessimistic
Else
    rs.Open "select * from pgr", con, adOpenKeyset, adLockPessimistic
    rs.AddNew
    rs("grp_idy") = increment("idy_PGR")
    MsgBox "New Category ID is " & rs("grp_idy"), vbExclamation, "DAS Version 4.0"
End If
rs("grp_cat") = Trim(Text2)
rs("grp_cls") = Trim(Text3)
rs.Update
clrctr M_FRM_PGR
Text1.Enabled = True
Text1.SetFocus
idnew = False
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub cancel()
clrctr M_FRM_PGR
Text1.Enabled = True
SSTab1.Tab = 0
Text1.SetFocus
idnew = False
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub newrec()
MOD1 = False
SSTab1.Tab = 0
clrctr M_FRM_PGR
Text1.Enabled = False
Text2.SetFocus
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub exitfrm()
Unload Me
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
ElseIf Check1.Value = 0 Then
    Check2.Value = 1
End If
prdgrplist
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
prdgrplist
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
If Text1.Enabled = True Then
Text1.SetFocus
Else
If Text2.Enabled = True Then Text2.SetFocus
End If
SSTab1.TabEnabled(1) = False
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_PGR.ActiveControl Is TextBox Then M_FRM_PGR.ActiveControl.Text = Trim(FUNKEY(K))
If list1.Visible = False Then
    If K = 78 And s = 2 Then
        newrec
    ElseIf K = 83 And s = 2 Then
        save
    ElseIf K = 81 And s = 2 Then
        cancel
    ElseIf K = 88 And s = 2 Then
        exitfrm
    ElseIf K = 68 And s = 2 Then
        delete
    ElseIf K = 67 And s = 2 Then
        If Toolbar1.Buttons(12).Enabled = True Then CREATE
    ElseIf K = 69 And s = 2 Then
        If M_FRM_PGR.ActiveControl.name = "Text1" Then prdgrplist
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And s = 2 Then
        OpenReport CURDIR & "\lists.RP1", "Product Categories list", 0, "", 6
    ElseIf K = 80 And s = 2 Then
        'MDIForm1.RRReport1.ReportName = "Categories list.RRW"
        'MDIForm1.RRReport1.Destination = 6
        'MDIForm1.RRReport1.RunReport 1
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        If CHECKTB("pgr.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            cancel
            Exit Sub
        End If

        If check("pgr", "grp_idy", Trim(Text1)) = False Then
            Text1.SetFocus
        Else
            Text1.Enabled = False
            Text2.SetFocus
        End If
    Else
        If M_FRM_PGR.ActiveControl.name = "Text1" Then
            Unload Me
        ElseIf M_FRM_PGR.ActiveControl.name = "SSTab1" Then
        Else
            save
        End If
    End If
End If
End Sub

Private Sub Form_Load()
checkdongle
MOD1 = False
idnew = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub List1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending1
End If
list1.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Product Group or Press <Esc> to Quit"
End Sub

Private Sub list1_KeyDown(KeyAscii As Integer, s As Integer)
Dim rs As New ADODB.Recordset
If KeyAscii = 13 Then
    If list1.ListItems.count > 0 Then
        If rs.State = 1 Then rs.Close
        Dim lst As ListSubItem
        Set lst = list1.SelectedItem.ListSubItems.Item(2)
        rs.Open "select * from pgr where grp_idy like '" & Trim(lst.Text) & "'", con, adOpenStatic
        If rs.RecordCount > 0 Then
            Text1 = Trim(rs("grp_idy"))
            Text2 = Trim(rs("grp_cat"))
            Text3 = Trim(rs("grp_cls") & "")
            list1.Visible = False
            SSTab1.Tab = 0
            Text1.Enabled = False
            Text2.SetFocus
            MOD1 = True
            Toolbar1.Buttons(12).Enabled = True
        End If
    End If
End If
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
Toolbar1.Enabled = True
list1.Visible = False
If SSTab1.Tab = 0 Then
    If Text1.Enabled = True Then
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Product Group ID or Press <Enter> to Get List of Product Group ID's"
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
End Sub
Private Sub text3_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Category"
End Sub

Private Sub Text3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Segment"
End Sub

Private Sub TEXT1_KeyDown(KeyCode As Integer, Shift As Integer)
If CHECKTB("pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
    If rs.State = 1 Then rs.Close
        rs.Open "select * from pgr where grp_idy like '" & Sincrement(Text1) & "'", con, adOpenStatic, adLockReadOnly
        If rs.RecordCount > 0 Then
        Text1 = Sincrement(Trim(Text1))
        Text2 = Trim(rs("grp_cat"))
        Text3 = Trim(rs("grp_cls"))
        Text1.Enabled = False
        Text2.SetFocus
        MOD1 = True
        Toolbar1.Buttons(12).Enabled = True
        Else
            prdgrplist
        End If
End If
End Sub

Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Text2)) > 0 Then
        Text3.SetFocus
    Else
        MsgBox "Category Should Not Be Blank, Enter Category", vbExclamation, "DAS Version 4.0"
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If CHECKTB("pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Text3)) > 0 Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from pgr where grp_cat like '" & Trim(Text2) & "' and grp_cls like '" & Trim(Text3) & "'", con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            If rs("grp_idy") = Trim(Text1) Then
                save
            Else
                MsgBox "Duplicate Values Are Invalid", vbExclamation, "DAS Version 4.0"
                Text3.SetFocus
            End If
        Else
            save
        End If
    Else
        MsgBox "Segment Should Not Be Blank, Enter Segment", vbExclamation, "DAS Version 4.0"
        Text3.SetFocus
    End If
ElseIf KeyCode = 38 Then
    If Len(Trim(Text3)) > 0 Then
            Text2.SetFocus
    Else
        MsgBox "Segment Should Not Be Blank, Enter Segment", vbExclamation, "DAS Version 4.0"
        Text3.SetFocus
    End If
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
    exitfrm
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "ed" Then
    If M_FRM_PGR.ActiveControl.name = "Text1" Then prdgrplist
ElseIf Button.KEY = "d" Then
    delete
ElseIf Button.KEY = "pre" Then
'    M_FRM_PRE.Caption = "Product Group ID Prefix"
'    M_FRM_PRE.Show 1
    SYS_PSW.Tag = "PGR_PRE"
    SYS_PSW.Show 1
ElseIf Button.KEY = "p" Then
    OpenReport CURDIR & "\lists.RP1", "Product Categories list", 0, "", 6
    cancel
ElseIf Button.KEY = "pr" Then
    Exit Sub
er1:
    MsgBox "Category Not Found, Define Category!", vbExclamation, "DAS Version 4.0"
ElseIf Button.KEY = "crt" Then
    CREATE
End If
End Sub
