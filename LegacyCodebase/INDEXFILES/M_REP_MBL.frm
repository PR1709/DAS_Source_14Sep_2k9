VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_REP_MBL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Option"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "M_REP_MBL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4471
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Report Filters"
      TabPicture(0)   =   "M_REP_MBL.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOk"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkAll"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_REP_MBL.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "lv"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox chkAll 
         Caption         =   "&All"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select"
         Height          =   720
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   3975
         Begin VB.TextBox txtMobile 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1440
            MaxLength       =   40
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mobile Number"
            Height          =   195
            Left            =   240
            TabIndex        =   3
            Top             =   315
            Width           =   1065
         End
      End
      Begin MSComctlLib.ListView lv 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   2778
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4440
         TabIndex        =   7
         Top             =   2715
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label9 
         Height          =   195
         Left            =   3120
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Number"
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
         Height          =   375
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   4180
      End
   End
End
Attribute VB_Name = "M_REP_MBL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAll_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If chkAll.Value = 1 Then
        cmdOk.SetFocus
    Else
        txtMobile.Enabled = True
        txtMobile.SetFocus
    End If
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If chkAll.Value = 1 Then
    txtMobile = ""
    OpenReport CURDIR & "\tcm.RP1", "Mobile Track Report", 2, "avt->mbl_idy <> ''", 6
Else
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    rs.Open "SELECT MBL_IDY FROM AVT WHERE MBL_IDY='" & Trim(txtMobile) & "'", con, adOpenStatic
    If rs.EOF Then
        MsgBox "Mobile ID not found!", vbInformation + vbOKOnly, "DAS Version 4.0"
        txtMobile.SetFocus
        SendKeys "{HOME}+{END}"
        Exit Sub
    Else
        OpenReport CURDIR & "\tcm.RP1", "Mobile Track Report", 2, "AVT->MBL_IDY='" & Trim(txtMobile) & "'", 6
    End If
End If
Unload Me
End Sub

Private Sub Form_Activate()
chkAll.SetFocus
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If SSTab1.Tab = 1 Then
        If Me.ActiveControl.name = "lv" Then
            SSTab1.Tab = 0
            chkAll.SetFocus
            SSTab1.TabEnabled(1) = False
        End If
    Else
        If Me.ActiveControl.name = "chkAll" Then
            Unload Me
        Else
            chkAll.SetFocus
            txtMobile = ""
        End If
    End If
End If
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

Private Sub lV_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If lv.ListItems.count > 0 Then
        txtMobile = lv.SelectedItem
        SSTab1.Tab = 0
        lv.Visible = False
        cmdOk.SetFocus
    End If
End If
End Sub

Private Sub txtMobile_KeyDown(K As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    Dim rsavt As New ADODB.Recordset
'    lv.ColumnHeaders.Clear
'    lv.ListItems.Clear
'    lv.ColumnHeaders.Add , , "Mobile NO", 2500
'    If Len(Trim(txtMobile)) > 0 Then
'        rsavt.Open "select mbl_idy from avt where mbl_idy like '" & Trim(txtMobile) & "%'", CON, adOpenStatic
'        If Not rsavt.EOF Then
'            cmdOk.SetFocus
'        Else
'            MsgBox "Mobile ID not found!", vbInformation + vbOKOnly, "DAS Version 4.0"
'            txtMobile.SetFocus
'            SendKeys "{home}+{end}"
'            Exit Sub
'        End If
'    Else
'        rsavt.Open "select mbl_idy from avt WHERE MBL_IDY IS NOT NULL OR MBL_IDY<>'' order by mbl_idy", CON, adOpenStatic
'        Dim ls As ListItem
'        While Not rsavt.EOF
'            Set ls = lv.ListItems.Add(, , Trim(rsavt(0)))
'        rsavt.MoveNext
'        Wend
'        If rsavt.RecordCount > 0 Then
'            SSTab1.TabEnabled(1) = True
'            SSTab1.Tab = 1
'            lv.Visible = True
'            lv.SetFocus
'        End If
'     End If
'End If
Dim ls As ListItem
If K = 13 Then
    Dim rs As New ADODB.Recordset
    rs.Open "select MBL_IDY from AVT where MBL_IDY like '" & Trim(txtMobile) & "%' and (mbl_idy is not null and mbl_idy<>'')", con, adOpenStatic
    lv.ListItems.CLEAR
    lv.ColumnHeaders.CLEAR
    If Not rs.EOF Then
        While Not rs.EOF
            Set ls = lv.ListItems.Add(, , Trim(rs("MBL_IDY")))
            rs.MoveNext
        Wend
        SSTab1.Tab = 1
        lv.ColumnHeaders.Add , , "Mobile Number"
        lv.Visible = True
        lv.SetFocus
'    If Not rs.EOF Then
'        txtMobile = rs("mbl_idy")
'        cmdOk.SetFocus
    Else
        MsgBox "Mobile ID not found!", vbInformation + vbOKOnly, "DAS Version 4.0"
        txtMobile.SetFocus
        SendKeys "{home}+{end}"
    End If
End If
End Sub

Private Sub txtMobile_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = 0
End Sub
Private Sub mblList()
Dim rs As New ADODB.Recordset
rs.Open "select mbl_idy from avt where mbl_idy<>'' and mbl_idy is not null order by mbl_idy", con, adOpenStatic
Dim ls As ListItem
lv.ListItems.CLEAR
    lv.ColumnHeaders.CLEAR
    If Not rs.EOF Then
        While Not rs.EOF
            Set ls = lv.ListItems.Add(, , Trim(rs("MBL_IDY")))
            rs.MoveNext
        Wend
        SSTab1.Tab = 1
        lv.ColumnHeaders.Add , , "Mobile Number"
        lv.Visible = True
        lv.SetFocus
    Else
        MsgBox "Mobile ID not found!", vbInformation + vbOKOnly, "DAS Version 4.0"
        txtMobile.SetFocus
        SendKeys "{home}+{end}"
    End If
End Sub
