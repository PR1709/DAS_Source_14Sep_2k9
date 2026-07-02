VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form USR_PSW 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Effmcg User Creation"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "USR_PSW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_PSW.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_PSW.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_PSW.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_PSW.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_PSW.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_PSW.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_PSW.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_PSW.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_PSW.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR_PSW.frx":269A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      DragMode        =   1  'Automatic
      Height          =   5195
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   430
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   9155
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      OLEDropMode     =   1
      TabCaption(0)   =   "User Creation"
      TabPicture(0)   =   "USR_PSW.frx":29EE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "USR_PSW.frx":2A0A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "list7"
      Tab(1).ControlCount=   2
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   660
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Check to enable module access"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   2775
         Left            =   120
         TabIndex        =   6
         Top             =   2100
         Width           =   9015
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   7560
            TabIndex        =   27
            Top             =   240
            Width           =   1300
            Begin VB.CheckBox Check7 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Reports"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   0
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   195
               Width           =   1215
            End
         End
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   4920
            TabIndex        =   25
            Top             =   240
            Width           =   1695
            Begin VB.CheckBox Check6 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Analysis"
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   0
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   70
               Width           =   975
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Caption         =   "Frame5"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3360
            TabIndex        =   23
            Top             =   240
            Width           =   1455
            Begin VB.CheckBox Check5 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Accounts"
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   120
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   75
               Width           =   1335
            End
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2040
            TabIndex        =   21
            Top             =   240
            Width           =   1215
            Begin VB.CheckBox Check4 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "Core"
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   0
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   70
               Width           =   975
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1455
            Begin VB.CheckBox Check3 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Caption         =   "File"
               ForeColor       =   &H80000008&
               Height          =   450
               Left            =   0
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   0
               Width           =   975
            End
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   1830
            ItemData        =   "USR_PSW.frx":2A26
            Left            =   240
            List            =   "USR_PSW.frx":2A45
            Style           =   1  'Checkbox
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   720
            Width           =   1575
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   1830
            ItemData        =   "USR_PSW.frx":2ACA
            Left            =   2040
            List            =   "USR_PSW.frx":2ADD
            OLEDropMode     =   1  'Manual
            Style           =   1  'Checkbox
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   720
            Width           =   1335
         End
         Begin VB.ListBox List3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   1830
            ItemData        =   "USR_PSW.frx":2B17
            Left            =   3480
            List            =   "USR_PSW.frx":2B2D
            Style           =   1  'Checkbox
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   720
            Width           =   1215
         End
         Begin VB.ListBox List4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   1830
            ItemData        =   "USR_PSW.frx":2B7C
            Left            =   4920
            List            =   "USR_PSW.frx":2B98
            Style           =   1  'Checkbox
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   720
            Width           =   2415
         End
         Begin VB.ListBox List5 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   1830
            ItemData        =   "USR_PSW.frx":2C47
            Left            =   7560
            List            =   "USR_PSW.frx":2C57
            Style           =   1  'Checkbox
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000A&
         Caption         =   "User Role"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   6240
         TabIndex        =   3
         Top             =   1140
         Width           =   2895
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "User"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1920
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Administrator"
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   240
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   360
            Width           =   1575
         End
      End
      Begin MSComctlLib.ListView list7 
         Height          =   3375
         Left            =   -74040
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1060
         Visible         =   0   'False
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5953
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Password"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "User Type"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select User"
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
         Left            =   -74040
         TabIndex        =   17
         Top             =   720
         Width           =   7215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   960
         TabIndex        =   16
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1065
         TabIndex        =   15
         Top             =   1140
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re-enter Password"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   405
         TabIndex        =   14
         Top             =   1620
         Width           =   1350
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
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
            Key             =   "s1"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "d"
            Object.ToolTipText     =   "Delete (Ctrl+D)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ed"
            Object.ToolTipText     =   "Browse (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "USR_PSW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mod1 As Boolean

Public Sub delete()
If Trim(Text1) <> "ADMINISTRATOR" Then
SYS_PSW.Caption = "User Deletion"
SYS_PSW.Show 1
If USR_PSW.Tag = "Delete" Then
    If MsgBox("Are you sure you want to delete User ?", vbYesNo + vbCritical, "Effmcg") = vbYes Then
        If check("usr", "usr_nme", Text1) = False Then
            MsgBox "User Not Found, Create User  !", vbExclamation, "Effmcg"
            SSTab1.Tab = 0
            clrctr USR_PSW
            Text1.Enabled = True
            Text1.SetFocus
            Exit Sub
        Else
            con.Execute "delete from usr where usr_nme like '" & Trim(Text1) & "'"
            SSTab1.Tab = 0
            Text1.Enabled = True
            Text1.SetFocus
            clrctr USR_PSW
            Option2.Value = True
        End If
        cancel
    Else
        cancel
    End If
End If
End If
End Sub

Private Sub dispdetails()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from usr where usr_nme like '" & Text1 & "'", con, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
    Text1.Enabled = False
    Text2 = Trim(rs("usr_pwd"))
    Text3 = Trim(rs("usr_pwd"))
    If Trim(rs("usr_typ")) = "Admin" Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
    For a = 1 To rs.RecordCount
        If Trim(rs("usr_pms")) = "File" Then
            CHECK3.Value = 1
        ElseIf Trim(rs("usr_pms")) = "Core" Then
            Check4.Value = 1
        ElseIf Trim(rs("usr_pms")) = "Accounts" Then
            Check5.Value = 1
        ElseIf Trim(rs("usr_pms")) = "Analysis" Then
            Check6.Value = 1
        ElseIf Trim(rs("usr_pms")) = "Reports" Then
            Check7.Value = 1
        Else
            For B = 0 To list1.ListCount - 1
                If Trim(rs("usr_pms")) = list1.List(B) Then
                list1.Selected(B) = True
                GoTo Next1
                End If
            Next
            For B = 0 To List2.ListCount - 1
                If Trim(rs("usr_pms")) = List2.List(B) Then
                List2.Selected(B) = True
                GoTo Next1
                End If
            Next
            For B = 0 To LIST3.ListCount - 1
                If Trim(rs("usr_pms")) = LIST3.List(B) Then
                LIST3.Selected(B) = True
                GoTo Next1
                End If
            Next
            For B = 0 To list4.ListCount - 1
                If Trim(rs("usr_pms")) = list4.List(B) Then
                list4.Selected(B) = True
                GoTo Next1
                End If
            Next
            For B = 0 To list5.ListCount - 1
                If Trim(rs("usr_pms")) = list5.List(B) Then
                list5.Selected(B) = True
                GoTo Next1
                End If
            Next
        End If
Next1:
rs.MoveNext
    Next
    SSTab1.Tab = 0
    Text2.SetFocus
    mod1 = True
Else
    If MsgBox("Creating New User...!", vbExclamation + vbYesNo, "Effmcg") = vbNo Then
        cancel
        Exit Sub
    Else
        Text2.SetFocus
    End If
End If
End Sub
Private Sub usrlist()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "Select distinct(usr_nme),usr_pwd,usr_typ from usr where usr_nme like '" & Trim(Text1) & "%'", con, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
    list7.ListItems.Clear
    For a = 1 To rs.RecordCount
        Set ls = list7.ListItems.Add(, , Trim(rs(0) & ""))
        ls.ListSubItems.Add , , Trim(rs(1))
        ls.ListSubItems.Add , , Trim(rs(2))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list7.Visible = True
    list7.SetFocus
Else
    MsgBox "User Not Found, Define User !", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    Text1.SetFocus
End If
End Sub
Private Sub Check3_Click()
If CHECK3.Value = 1 Then
Frame1.Enabled = True
list1.Enabled = True
list1.SetFocus
Else
list1.Enabled = False
End If
End Sub

Private Sub Check3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check File To Enable File Menu List"
End Sub

Private Sub Check4_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Core To Enable Core Menu List"
End Sub

Private Sub Check5_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Accounts To Enable Accounts Menu List"
End Sub

Private Sub Check6_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check MIR's To Enable MIR's Menu List"
End Sub

Private Sub Check7_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Reports To Enable Reports Menu List"
End Sub

Private Sub Check3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And CHECK3.Value = 1 Then
    list1.SetFocus
Else
    Check4.SetFocus
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
Frame1.Enabled = True
List2.Enabled = True
List2.SetFocus
Else
List2.Enabled = False
End If
End Sub

Private Sub Check4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Check4.Value = 1 Then
    List2.SetFocus
Else
    Check5.SetFocus
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
Frame1.Enabled = True
LIST3.Enabled = True
LIST3.SetFocus
Else
LIST3.Enabled = False
End If
End Sub

Private Sub Check5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Check5.Value = 1 Then
    LIST3.SetFocus
Else
    Check6.SetFocus
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
Frame1.Enabled = True
list4.Enabled = True
list4.SetFocus
Else
list4.Enabled = False
End If
End Sub

Private Sub Check6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Check6.Value = 1 Then
    list4.SetFocus
Else
    Check7.SetFocus
End If
End Sub

Private Sub Check7_CLICK()
If Check7.Value = 1 Then
Frame1.Enabled = True
list5.Enabled = True
list5.SetFocus
Else
'Frame1.Enabled = False
list5.Enabled = False
End If
End Sub

Private Sub Check7_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 13 And Check7.Value = 1) Or (KeyCode = 32 And Check7.Value = 0) Then
    Frame1.Enabled = True
    list5.Enabled = True
    list5.SetFocus
Else
    Check7.Value = 0
    save
End If
End Sub

Private Sub save()

Dim rs As New ADODB.Recordset
If MsgBox("Confirm Record Save ?", vbExclamation + vbYesNo, "Effmcg") = vbNo Then
    cancel
    Exit Sub
End If

If Len(Trim(Text1)) = 0 Then
MsgBox "Username Should Not Be Blank, Define Username !", vbExclamation, "Effmcg"
SSTab1.Tab = 0
Text1.SetFocus
Exit Sub
End If

If Len(Trim(Text2)) < 4 Then
MsgBox "Password Should Not be Less Than Four Characters !", vbExclamation, "Effmcg"
Text2 = ""
Text3 = ""
SSTab1.Tab = 0
Text2.SetFocus
Exit Sub
End If

If Trim(Text3) <> Trim(Text2) Then
MsgBox "Password Doesn't Match, Re-enter Password !", vbExclamation, "Effmcg"
SSTab1.Tab = 0
Text3.SetFocus
Exit Sub
End If

If mod1 = False Then
Else
con.Execute "delete from usr where usr_nme like '" & Text1 & "'"
End If

If rs.State = 1 Then rs.Close
If Trim(Text1) = "ADMINISTRATOR" Then
    rs.Open "select * from usr", con, adOpenKeyset, adLockOptimistic
    rs.AddNew
    rs("usr_nme") = Trim(Text1)
    rs("usr_pwd") = Trim(Text2)
    rs("usr_typ") = "Admin"
    rs.Update
    clrctr USR_PSW
    Text1.Enabled = True
    Text1.SetFocus
    Exit Sub
End If
If rs.State = 1 Then rs.Close
rs.Open "select * from usr", con, adOpenKeyset, adLockOptimistic
If Option1.Value = True Then
    rs.AddNew
    rs("usr_nme") = Trim(Text1)
    rs("usr_pwd") = Trim(Text2)
    rs("usr_typ") = "Admin"
    rs.Update
    clrctr USR_PSW
    Text1.Enabled = True
    Text1.SetFocus
    Exit Sub
End If

If CHECK3.Value = 1 Then
        rs.AddNew
        rs("usr_nme") = Trim(Text1)
        rs("usr_pwd") = Trim(Text2)
        rs("usr_pms") = CHECK3.Caption
        rs("usr_typ") = "User"
        rs.Update
For a = 0 To list1.ListCount - 1
    If list1.Selected(a) = True Then
        rs.AddNew
        rs("usr_nme") = Trim(Text1)
        rs("usr_pwd") = Trim(Text2)
        rs("usr_pms") = list1.List(a)
        rs("usr_typ") = "User"
        rs.Update
    End If
Next
End If
If Check4.Value = 1 Then
        rs.AddNew
        rs("usr_nme") = Trim(Text1)
        rs("usr_pwd") = Trim(Text2)
        rs("usr_pms") = Check4.Caption
        rs("usr_typ") = "User"
        rs.Update
For a = 0 To List2.ListCount - 1
    If List2.Selected(a) = True Then
        rs.AddNew
        rs("usr_nme") = Trim(Text1)
        rs("usr_pwd") = Trim(Text2)
        rs("usr_pms") = List2.List(a)
        rs("usr_typ") = "User"
        rs.Update
    End If
Next
End If

If Check5.Value = 1 Then
        rs.AddNew
        rs("usr_nme") = Trim(Text1)
        rs("usr_pwd") = Trim(Text2)
        rs("usr_pms") = Check5.Caption
        rs("usr_typ") = "User"
        rs.Update
For a = 0 To LIST3.ListCount - 1
    If LIST3.Selected(a) = True Then
        rs.AddNew
        rs("usr_nme") = Trim(Text1)
        rs("usr_pwd") = Trim(Text2)
        rs("usr_pms") = LIST3.List(a)
        rs("usr_typ") = "User"
        rs.Update
    End If
Next
End If

If Check6.Value = 1 Then
        rs.AddNew
        rs("usr_nme") = Trim(Text1)
        rs("usr_pwd") = Trim(Text2)
        rs("usr_pms") = Check6.Caption
        rs("usr_typ") = "User"
        rs.Update
For a = 0 To list4.ListCount - 1
    If list4.Selected(a) = True Then
        rs.AddNew
        rs("usr_nme") = Trim(Text1)
        rs("usr_pwd") = Trim(Text2)
        rs("usr_pms") = list4.List(a)
        rs("usr_typ") = "User"
        rs.Update
    End If
Next
End If

If Check7.Value = 1 Then
        rs.AddNew
        rs("usr_nme") = Trim(Text1)
        rs("usr_pwd") = Trim(Text2)
        rs("usr_pms") = Check7.Caption
        rs("usr_typ") = "User"
        rs.Update
For a = 0 To list5.ListCount - 1
    If list5.Selected(a) = True Then
        rs.AddNew
        rs("usr_nme") = Trim(Text1)
        rs("usr_pwd") = Trim(Text2)
        rs("usr_pms") = list5.List(a)
        rs("usr_typ") = "User"
        rs.Update
    End If
Next
End If
clrctr USR_PSW
For a = 1 To USR_PSW.Count - 1
If TypeOf USR_PSW.Controls(a) Is ListBox Then
    For B = 0 To USR_PSW.Controls(a).ListCount - 1
        USR_PSW.Controls(a).Selected(B) = False
    Next
ElseIf TypeOf USR_PSW.Controls(a) Is CheckBox Then
    USR_PSW.Controls(a).Value = 0
End If
Next
SSTab1.Tab = 0
Text1.Enabled = True
Text1.SetFocus
Frame1.Enabled = False
mod1 = False
Option2.Value = True
End Sub

Private Sub cancel()
clrctr USR_PSW
For a = 1 To USR_PSW.Count - 1
If TypeOf USR_PSW.Controls(a) Is ListBox Then
    For B = 0 To USR_PSW.Controls(a).ListCount - 1
        USR_PSW.Controls(a).Selected(B) = False
    Next
ElseIf TypeOf USR_PSW.Controls(a) Is CheckBox Then
    USR_PSW.Controls(a).Value = 0
End If
Next
SSTab1.Tab = 0
Text1.Enabled = True
Text1.SetFocus
Frame1.Enabled = False
mod1 = False
Option2.Value = True
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub newrec()
Dim rs As New ADODB.Recordset
cancel
rs.Open "select * from usr", con, adOpenKeyset, adLockOptimistic
rs.AddNew
mod1 = False
Option2.Value = True
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub exitfrm()
Unload Me
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
If Text1.Enabled = True Then
    Text1.SetFocus
End If
SSTab1.TabEnabled(1) = False
End Sub

Private Sub Form_KeyDown(K As Integer, S As Integer)
If K = 120 Then Shell "c:\windows\calc.exe"
If K > 111 And K < 120 Then If TypeOf USR_PSW.ActiveControl Is TextBox Then USR_PSW.ActiveControl.Text = Trim(FUNKEY(K))
If list7.Visible = False Then
    If K = 78 And S = 2 Then
        newrec
    ElseIf K = 83 And S = 2 Then
        save
    ElseIf K = 81 And S = 2 Then
        cancel
    ElseIf K = 88 And S = 2 Then
        exitfrm
    ElseIf K = 68 And S = 2 Then
        delete
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 69 And S = 2 Then
        If USR_PSW.ActiveControl.name = "Text1" Then usrlist
    ElseIf K = 80 And S = 2 Then
        'MDIForm1.RRReport1.ReportName = App.Path & "\users list.rrw"
        On Error GoTo er
        'MDIForm1.RRReport1.Destination = 2
        'MDIForm1.RRReport1.Printer = "?"
        'MDIForm1.RRReport1.RunReport 1
        Exit Sub
er:
        MsgBox "Printer Unavailable, Set Printer !", vbExclamation, "Effmcg"
    ElseIf K = 82 And S = 2 Then
    On Error GoTo er1
        'MDIForm1.RRReport1.ReportName = App.Path & "\users list.rrw"
        'MDIForm1.RRReport1.Destination = 6
        'MDIForm1.RRReport1.RunReport 1
        Exit Sub
er1:
    MsgBox "Users Not Found, Define User...!", vbExclamation, "Effmcg"
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list7.Visible = True Then
        list7.Visible = False
        SSTab1.Tab = 0
        If Text1.Enabled = True Then
            Text1.SetFocus
        Else
            Text2.SetFocus
        End If
    Else
        If USR_PSW.ActiveControl.name = "Text1" Then
            Unload Me
        Else
            save
        End If
    End If
End If
End Sub

Private Sub Form_Load()
Option2.Value = True
mod1 = False
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub list1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Items In the List To Enable Access Permission"
End Sub

Private Sub list2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Items In the List To Enable Access Permission"
End Sub

Private Sub list3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Items In the List To Enable Access Permission"
End Sub

Private Sub list4_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Items In the List To Enable Access Permission"
End Sub

Private Sub List5_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Check Items In the List To Enable Access Permission"
End Sub

Private Sub list1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Check4.SetFocus
End Sub

Private Sub List2_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Check5.SetFocus
End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Check6.SetFocus
End Sub

Private Sub List4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Check7.SetFocus
End Sub

Private Sub List5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then save
End Sub

Private Sub list7_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list7.SortOrder = lvwAscending Then
list7.SortOrder = lvwDescending
ElseIf list7.SortOrder = lvwDescending Then
list7.SortOrder = lvwAscending
End If
list7.SortKey = CH.Index - 1
End Sub

Private Sub list7_DblClick()
List7_Keydown 13, 0
End Sub

Private Sub List7_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select User or Press <Esc> to Quit"
End Sub

Private Sub List7_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text1 = list7.SelectedItem
    dispdetails
    SSTab1.Tab = 0
    Text2.SetFocus
    list7.Visible = False
End If
End Sub

Private Sub List7_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub Option1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> For Creating Administrator"
End Sub

Private Sub Option2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> For Creating User"
End Sub

Private Sub Option1_KeyPress(K As Integer)
If K = 13 Then
If Option1.Value = True Then
Frame1.Enabled = False
save
End If
End If
End Sub

Private Sub Option2_KeyPress(K As Integer)
If K = 13 Then
If Option2.Value = True Then
Frame1.Enabled = True
CHECK3.SetFocus
End If
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If list7.Visible = True Then list7.Visible = False
Toolbar1.Enabled = True
If SSTab1.Tab = 0 Then
If Text1.Enabled = True Then
Text1.SetFocus
Else
Text2.SetFocus
End If
End If
End Sub
 
Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter User Name"
End Sub

Private Sub Text1_KeyDown(K As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Then
    If Len(Trim(Text1)) > 0 Then
        dispdetails
    Else
        usrlist
    End If
End If
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Password"
End Sub

Private Sub Text2_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    If Len(Trim(Text2)) >= 4 Then
        Text3.SetFocus
    Else
        MsgBox "Password Should Not be Less Than Four Characters !", vbExclamation, "Effmcg"
        Text2 = ""
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Text3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Re-enter Password"
End Sub

Private Sub Text3_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
    If Trim(Text3) = Trim(Text2) Then
        If Option1.Value = True Then
            Option1.SetFocus
        Else
            Option2.SetFocus
        End If
    Else
        MsgBox "Password Doesn't Match, Re-enter Password !", vbExclamation, "Effmcg"
        Text3 = ""
        Text3.SetFocus
    End If
ElseIf K = 38 Then
    Text2.SetFocus
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
    If USR_PSW.ActiveControl.name = "Text1" Then usrlist
ElseIf Button.KEY = "d" Then
    delete
    
ElseIf Button.KEY = "p" Then
    'MDIForm1.RRReport1.ReportName = App.Path & "\Users list.rrw"
    'On Error GoTo er
    'MDIForm1.RRReport1.Destination = 2
    'MDIForm1.RRReport1.Printer = "?"
    'MDIForm1.RRReport1.RunReport 1
    'Exit Sub
'er:
 '   MsgBox "Printer Unavailable, Set Printer !", vbExclamation, "Effmcg"
ElseIf Button.KEY = "pr" Then
    'On Error GoTo er1
    'MDIForm1.RRReport1.ReportName = App.Path & "\Users list.rrw"
    'MDIForm1.RRReport1.Destination = 6
    'MDIForm1.RRReport1.RunReport 1
    'Exit Sub
'er1: '
 '   MsgBox "Users Not Found, Define User...!", vbExclamation, "Effmcg"
End If
End Sub
