VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_CAT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retailers' Categories - Channels"
   ClientHeight    =   5625
   ClientLeft      =   -330
   ClientTop       =   1440
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "M_FRM_CAT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
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
            Picture         =   "M_FRM_CAT.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAT.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAT.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAT.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAT.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAT.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAT.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAT.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAT.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAT.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAT.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAT.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5140
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9075
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
      TabCaption(0)   =   "&Categories - Channel Details"
      TabPicture(0)   =   "M_FRM_CAT.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_CAT.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "list1"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
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
         Left            =   2100
         MaxLength       =   35
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1440
         Width           =   3945
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
         Left            =   2100
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1065
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Category Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Retailers' Category"
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
         TabIndex        =   6
         Top             =   720
         Width           =   7900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category - Channel ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   375
         TabIndex        =   5
         Top             =   1035
         Width           =   1560
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category - Channel Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1515
         Width           =   1815
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
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
            Object.ToolTipText     =   "Prefix Area ID"
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
Attribute VB_Name = "M_FRM_CAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MOD1 As Boolean, idnew As Boolean

Private Sub CREATE()
MOD1 = False
Text1 = ""
save
End Sub

Public Sub delete()
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If MsgBox("Are you sure you want to delete Category ?", vbYesNo + vbCritical, "DAS Version 4.0") = vbYes Then
    If check("CAT", "CAT_idy", Text1) = False Then
        MsgBox "Category Not Found, Define Category !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        clrctr M_FRM_CAT
        Text1.Enabled = True
        Text1.SetFocus
        Exit Sub
    Else
        If deletecheck = False Then
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
                    con.Execute "delete from CAT where CAT_idy like '" & Sincrement(Text1) & "'"
                    SSTab1.Tab = 0
                    Text1.Enabled = True
                    Text1.SetFocus
                    clrctr M_FRM_CAT
                End If
            End If
            '
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
Private Function deletecheck() As Boolean
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Function
End If
If check("act", "CAT_idy", Trim(Text1)) = False Then
    deletecheck = False
Else
    deletecheck = True
End If
End Function
Private Sub save()
Dim rs As New ADODB.Recordset
If MsgBox("Confirm Record Save ?", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
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

If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If Len(Trim(Text2)) = 0 Then
    MsgBox "Category Name Should Not Be Blank, Give Name !", vbExclamation, "DAS Version 4.0"
    Text2.Enabled = True
    Text2.SetFocus
    Exit Sub
End If
If rs.State = 1 Then rs.Close
If MOD1 = True Then
rs.Open "select * from CAT where CAT_idy like '" & Sincrement(Text1) & "'", con, adOpenKeyset, adLockPessimistic
Else
rs.Open "select * from CAT", con, adOpenKeyset, adLockPessimistic
rs.AddNew
rs("CAT_idy") = increment("idy_cat")
MsgBox "New Category ID is " & rs("CAT_idy"), vbExclamation, "DAS Version 4.0"
End If
rs("CAT_nme") = Trim(Text2)
rs.Update
clrctr M_FRM_CAT
Text1.Enabled = True
idnew = False
SSTab1.Tab = 0

Text1.SetFocus
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub cancel()
clrctr M_FRM_CAT
Text1.Enabled = True
idnew = False
SSTab1.Tab = 0
Text1.SetFocus
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub newrec()
MOD1 = False
clrctr M_FRM_CAT
Text1.Enabled = False
SSTab1.Tab = 0
Text2.SetFocus
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub channellist()
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from CAT where CAT_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("CAT_nme")))
        ls.ListSubItems.Add , , Trim(rs("CAT_idy"))
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

Private Sub exitfrm()
Unload Me
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
If Text1.Enabled = True Then
    Text1.SetFocus
Else
    If Text2.Enabled = True Then Text2.SetFocus
End If
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf M_FRM_CAT.ActiveControl Is TextBox Then M_FRM_CAT.ActiveControl.Text = Trim(FUNKEY(K))
If list1.Visible = False Then
    If K = 78 And s = 2 Then
        newrec
    ElseIf K = 83 And s = 2 Then
        save
    ElseIf K = 81 And s = 2 Then
        cancel
    ElseIf K = 88 And s = 2 Then
        exitfrm
    ElseIf K = 67 And s = 2 Then
        If Toolbar1.Buttons(12).Enabled = True Then CREATE
    ElseIf K = 68 And s = 2 Then
        delete
    ElseIf K = 69 And s = 2 Then
        If M_FRM_CAT.ActiveControl.name = "Text1" Then channellist
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And s = 2 Then
        OpenReport CURDIR & "\lists.RP1", "Categories list", 0, "", 6
    ElseIf K = 82 And s = 2 Then
        'MDIForm1.RRReport1.ReportName = App.Path & "\Channels list.rrw"
        'MDIForm1.RRReport1.Destination = 1
        'MDIForm1.RRReport1.Printer = "?"
        'MDIForm1.RRReport1.RunReport 1
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        SSTab1.TabEnabled(1) = False
        If CHECKTB("cat.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            cancel
            Exit Sub
        End If
        If check("CAT", "CAT_idy", Trim(Text1)) = False Then
            Text1.Enabled = True
            Text1.SetFocus
        Else
            Text1.Enabled = False
            Text2.SetFocus
        End If
    Else
        If M_FRM_CAT.ActiveControl.name = "Text1" Then
            Unload Me
        ElseIf M_FRM_CAT.ActiveControl.name = "SSTab1" Then
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

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Category or Press <Esc> to Quit"
End Sub

Private Sub list1_KeyDown(K As Integer, s As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Then
    If list1.ListItems.count > 0 Then
        Dim lst As ListSubItem
        Set lst = list1.SelectedItem.ListSubItems.Item(1)
        If rs.State = 1 Then rs.Close
        rs.Open "select * from CAT where CAT_idy like '" & Trim(lst.Text) & "'", con, adOpenStatic
        Text1 = Trim(rs("CAT_idy"))
        Text2 = Trim(rs("CAT_nme"))
        list1.Visible = False
        SSTab1.TabEnabled(1) = False
        SSTab1.Tab = 0
        Text1.Enabled = False
        Text2.SetFocus
        Toolbar1.Buttons(12).Enabled = True
        MOD1 = True
    End If
End If
End Sub

Private Sub SSTab1_DblClick()
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
M_MDI.STS_BAR.Panels(1).Text = "Type Category ID or Press <Enter> to get List of Category ID's"
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 27 And Text1 = "" Then Unload Me
End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Category Name"
End Sub

Private Sub TEXT1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If rs.State = 1 Then rs.Close
        rs.Open " select * from cAT where cAT_idy like '" & Sincrement(Text1) & "'", con, adOpenStatic
        If rs.RecordCount > 0 Then
            Text1 = Sincrement(Text1)
            Text2 = Trim(rs("cAT_nme"))
            Text1.Enabled = False
            Text2.SetFocus
            Toolbar1.Buttons(12).Enabled = True
            MOD1 = True
        Else
            channellist
        End If
End If
End Sub

Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Text2)) > 0 Then
       save
     Else
        MsgBox "Category Name Should Not Be Blank, Give Name !", vbExclamation, "DAS Version 4.0"
        Text2.SetFocus
     End If
End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
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
    If M_FRM_CAT.ActiveControl.name = "Text1" Then channellist
ElseIf Button.KEY = "d" Then
    delete
ElseIf Button.KEY = "pre" Then
    SYS_PSW.Tag = "CAT_PRE"
    SYS_PSW.Show 1
ElseIf Button.KEY = "p" Then
    OpenReport CURDIR & "\lists.RP1", "Categories list", 0, "", 6
    cancel
ElseIf Button.KEY = "pr" Then
    'MDIForm1.RRReport1.WindowTitle = "Retailers' Categories List Report"
    'MDIForm1.RRReport1.ReportName = App.Path & "\Channels list.rrw"
    'MDIForm1.RRReport1.Destination = 1
    'MDIForm1.RRReport1.Printer = "?"
    'MDIForm1.RRReport1.RunReport 1
ElseIf Button.KEY = "crt" Then
    CREATE
End If
End Sub
