VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form M_FRM_HHT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Mobile Details"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "M_FRM_HHT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleMode       =   0  'User
   ScaleWidth      =   8820
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
            Picture         =   "M_FRM_HHT.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HHT.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HHT.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HHT.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HHT.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HHT.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HHT.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HHT.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HHT.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HHT.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HHT.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_HHT.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5205
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   465
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9181
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
      TabCaption(0)   =   "&Hand Held Details"
      TabPicture(0)   =   "M_FRM_HHT.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt_inc"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command10"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command11"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "MSComm1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_HHT.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "list1"
      Tab(1).ControlCount=   2
      Begin MSCommLib.MSComm MSComm1 
         Left            =   6660
         Top             =   1035
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Send to HH"
         Height          =   375
         Left            =   6960
         TabIndex        =   26
         Top             =   3390
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Import Bills"
         Height          =   375
         Left            =   6945
         TabIndex        =   25
         Top             =   4275
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   6000
         TabIndex        =   24
         Text            =   "Text6"
         Top             =   3240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   6000
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   3120
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6000
         TabIndex        =   21
         Top             =   3240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6960
         TabIndex        =   20
         Top             =   3810
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "C&reate File"
         Height          =   375
         Left            =   6975
         TabIndex        =   19
         Top             =   2955
         Width           =   1455
      End
      Begin VB.TextBox txt_inc 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Frame Frame1 
         Caption         =   "Hand Held Bill Header Print Details"
         Height          =   2295
         Left            =   720
         TabIndex        =   8
         Top             =   2400
         Width           =   5055
         Begin VB.TextBox txt_tfx 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1800
            Width           =   3345
         End
         Begin VB.TextBox txt_ad2 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1320
            Width           =   3345
         End
         Begin VB.TextBox txt_ad1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   840
            Width           =   3345
         End
         Begin VB.TextBox txt_dis 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1185
            MaxLength       =   20
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   375
            Width           =   3345
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Others"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   465
            TabIndex        =   18
            Top             =   1868
            Width           =   465
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Others"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   465
            TabIndex        =   17
            Top             =   1388
            Width           =   465
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   915
            Width           =   570
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   510
            TabIndex        =   10
            Top             =   435
            Width           =   420
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1860
         MaxLength       =   20
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1455
         Width           =   3945
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   3
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
            Text            =   "MIBS Salesman Description"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "* Click on Create File to Generate Bill Header Information"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   27
         Top             =   4800
         Width           =   3990
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Start Number"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   1995
         Width           =   1170
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   855
         TabIndex        =   6
         Top             =   1515
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ID"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   1035
         Width           =   210
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Mobile Instant Billing Source ID and Reference"
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
         TabIndex        =   4
         Top             =   720
         Width           =   7900
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Area ID"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   4560
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "M_FRM_HHT"
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
If CHECKTB("HHT.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If MsgBox("Are you sure you want to delete Hand Held ID ?", vbYesNo + vbCritical, "DAS Version 4.0") = vbYes Then
    If check("HHT", "HHT_IDY", Text1) = False Then
        MsgBox "Hand Held ID Not Found, Define Hand Held!", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        clrctr M_FRM_HHT
        Text1.Enabled = True
        Text1.SetFocus
        Exit Sub
    Else
        If deletecheck = False Then
            con.Execute "delete from HHT where HHT_IDY like '" & (Text1) & "'"
            SSTab1.Tab = 0
            Text1.Enabled = True
            Text1.SetFocus
            clrctr M_FRM_HHT
        Else
            MsgBox "Cannot Delete, Transaction Exists!", vbExclamation, "DAS Version 4.0"
            cancel
            Text1.SetFocus
        End If
    End If
    cancel
Else
    cancel
End If
End Sub
Public Function deletecheck() As Boolean
If CHECKTB("HHT.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Function
End If
If check("HHT", "HHT_IDY", Trim(Text1)) = False Then
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
If CHECKTB("HHT.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If Len(Trim(Text2)) = 0 Then
    MsgBox "Hand Held Details Should Not Be Blank!", vbExclamation, "DAS Version 4.0"
    Text2.Enabled = True
    Text2.SetFocus
    Exit Sub
End If
If rs.State = 1 Then rs.Close
If MOD1 = True Then
rs.Open "select * from HHT where HHT_IDY like '" & (Text1) & "'", PNL, adOpenKeyset, adLockPessimistic
Else
    Dim hht_id As String
    If rs.State = 1 Then rs.Close
        rs.Open "select max(HHT_IDY) from HHT", PNL, adOpenKeyset, adLockPessimistic
    If rs.RecordCount = 0 Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from HHT", PNL, adOpenKeyset, adLockPessimistic
        rs.AddNew
        rs("HHT_IDY") = "0000000001"
    Else
        hht_id = Mid(rs(0), 4, Len(rs(0)))
        hht_id = hht_id + 1
        While (Len(hht_id) < 10)
        hht_id = "0" & hht_id
        Wend
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM HHT", PNL, adOpenKeyset, adLockPessimistic
        rs.AddNew
        Text2.Locked = False
        rs("HHT_IDY") = hht_id
    End If
    Dim F2C As String
    MsgBox "New Hand Held ID is " & rs("HHT_IDY"), vbExclamation, "DAS Version 4.0"
    
End If
Dim PS, ls As String
ls = Mid(rs("HHT_IDY"), 9)
Text2.Text = Trim(Text2.Text)
PS = Mid(Text2.Text, 1, 2)
TEXTLENGTH = Len(txt_inc.Text)
If TEXTLENGTH = 1 Then
txt_inc.Text = "00000" & txt_inc.Text
ElseIf TEXTLENGTH = 2 Then
txt_inc.Text = "0000" & txt_inc.Text
ElseIf TEXTLENGTH = 3 Then
txt_inc.Text = "000" & txt_inc.Text
ElseIf TEXTLENGTH = 4 Then
txt_inc.Text = "00" & txt_inc.Text
ElseIf TEXTLENGTH = 5 Then
txt_inc.Text = "0" & txt_inc.Text
ElseIf TEXTLENGTH = 6 Then
txt_inc.Text = txt_inc.Text
End If
txt_inc.Text = PS & ls & txt_inc.Text
rs("HHT_NME") = Trim(Text2)
If MOD1 = False Then
rs("HHT_inc") = txt_inc.Text
rs("HHT_cur") = "1" 'Trim(txt_inc)
End If
rs("dst_NME") = Trim(txt_dis)
rs("dst_ad1") = Trim(txt_ad1)
rs("DST_ad2") = Trim(txt_ad2)
rs("DST_TFX") = Trim(txt_tfx)
rs.Update
clrctr M_FRM_HHT
Text1.Enabled = True
idnew = False
SSTab1.Tab = 0
Text1.SetFocus
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub cancel()
clrctr M_FRM_HHT
Text1.Enabled = True
idnew = False
SSTab1.Tab = 0
Text1.SetFocus
MOD1 = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub newrec()
UNLOCKCONTROL
MOD1 = False
clrctr M_FRM_HHT
Text1.Enabled = False
SSTab1.Tab = 0
Text2.SetFocus
Text2.Locked = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub arealist()
If CHECKTB("HHT.DBF") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from HHT where HHT_NME like '" & Trim(Text1) & "%'", PNL, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("HHT_NME")))
        ls.ListSubItems.Add , , Trim(rs("HHT_IDY"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.ListItems(1).Selected = True
    list1.SetFocus
Else
    MsgBox "Hand Held ID Not Found, Define Hand Held!", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub exitfrm()
Unload Me
End Sub
Private Sub Command1_Click()
Dim INL_CHK As New ADODB.Recordset
Dim s As Integer
Dim ss As String
Dim f
If MsgBox("Create Bill Header Information for Hand Held?", vbYesNo, "DAS Version 4.0") = vbYes Then
    If INL_CHK.State = 1 Then INL_CHK.Close
    INL_CHK.Open "SELECT * FROM INL where INV_IDY = '" & Trim(txt_inc) & "' ", con, adOpenKeyset, adLockPessimistic
    If INL_CHK.RecordCount > 0 Then
        MsgBox "Create a New Hand Held ID! Bills Exist with this ID!", vbCritical, "DAS Version 4.0"
        Exit Sub
    End If
    Dim FS_SOU As Object
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
        FS_SOU.CreateFolder (CURDIR & "\HHT")
    End If
    If FS_SOU.FileExists(CURDIR & "\HHT\BHD.TXT") = True Then
        FS_SOU.DeleteFile (CURDIR & "\HHT\BHD.TXT")
    End If
    Set f = FS_SOU.OpenTextFile(CURDIR & "\HHT\BHD.TXT", 2, True)
    NME_LEN = 40 - Len(txt_dis)
    AD1_LEN = 40 - Len(txt_ad1)
    AD2_LEN = 40 - Len(txt_ad2)
    TFX_LEN = 20 - Len(txt_tfx)
    HHT_LEN = 20 - Len(Text2.Text)
    DST_NME = txt_dis & Space(NME_LEN)
    DST_AD1 = txt_ad1 & Space(AD1_LEN)
    DST_AD2 = txt_ad2 & Space(AD2_LEN)
    DST_TFX = txt_tfx & Space(TFX_LEN)
    HHT_NME = Text2.Text & Space(HHT_LEN)
    f.WriteLine DST_NME & DST_AD1 & DST_AD2 & DST_TFX & Text1.Text & HHT_NME & txt_inc
    f.Close
    If Len(Trim(Text1)) = 10 And Len(Trim(Text2)) > 0 Then
    MsgBox "Bill Header Information Created for Hand Held Export", vbInformation, "DAS Version 4.0"
    Command11.Enabled = True
    Else
    MsgBox "Hand Held Identification  not selected for Hand Held Export!"
'    Text1.SetFocus
    End If
    End If
    cancel
    
End Sub
Private Sub Command10_Click()
Command10.Enabled = False
Dim RS1 As New ADODB.Recordset
Dim ACT_CHK As New ADODB.Recordset
Dim s, EPRD_UOM, ETOT_PCS, i, ENET_AMT, EMRP_RAT, NRS, TOTREC, CONFAC, SUBUNT As Integer
Dim ss, INL, STKCHECK, str As String
Dim f
Dim LINV_IDY, LCST_IDY, LGRS_AMT, LCSD_AMT, LNET_AMT, LSMN_IDY, LTXN_TYP, LROU_IDY, LUSR_IDY, LSYS_DAT, LSYS_TIM, LSRN_AMT, EINV_IDY, EINV_DAT, EPRD_IDY, EPRD_QTY, EPRD_RAT, ERATAPP, ECOL_AMT, EORD_IDY, ETXN_TYP, EFST_FLG, EBAT_IDY, ESUB_UNT As String, ILDS_IDY
Dim rs As New ADODB.Recordset
Dim ERSI_IDY, EvAN_CLS, ECAS_AMT, ECHQ_AMT, EDEP_RCD, EDEP_REF As String
Dim PRS As New ADODB.Recordset
Dim LOGRS As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim ARS As New ADODB.Recordset
Dim PRS1 As New ADODB.Recordset
Dim PGRRS As New ADODB.Recordset
Dim BRS As New ADODB.Recordset
Dim PNLRS As New ADODB.Recordset
Dim ACTRS As New ADODB.Recordset
Dim IRTRS As New ADODB.Recordset
'' End of Record Set Declarations
   Dim LOCRS As New ADODB.Recordset
   Dim MARRS As New ADODB.Recordset
   Dim CATRS As New ADODB.Recordset
   Dim SMNRS As New ADODB.Recordset
   Dim DIRRS As New ADODB.Recordset
   Dim DIRRS1 As New ADODB.Recordset
   Dim DIRRS2 As New ADODB.Recordset
   Dim DIRRS3 As New ADODB.Recordset
   Dim DIRRS4 As New ADODB.Recordset
   Dim DIRRS5 As New ADODB.Recordset
   Dim DIRRS6 As New ADODB.Recordset
   Dim DIRRS7 As New ADODB.Recordset
   Dim DIRRS8 As New ADODB.Recordset
   Dim DIRRS9 As New ADODB.Recordset
   
Dim ECST_NME, ECST_AD1, ECST_AD2, ECST_AD3, ECSTAD4, ECST_CST, ECST_GST
Dim EPRD_NME, EBRD_NME, EGRP_IDY, EGRP_CLS, EGRP_CAT, ELST_RAT, ELST_DIS, EBAT_NME As String

' File Total Records, File Name and Input Data Declarations
Dim TOT_REC As Integer
Dim FILE_NAME As String
Dim input_data As Variant
' End of File Declarations
If CHECKTB("INL.DBF,INV.DBF,PNL.DBF,PNV.DBF,SNV.DBF,VAN.DBF,IRT.DBF,ACT.DBF,LOG.DBF") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & "DAS Version 4.0"
    Exit Sub
End If
Dim TOTCASAMT, TOTCHQAMT, TOTDEPRCD, TOTDEPREF As Integer
TOTCASAMT = 0
TOTCHQAMT = 0
TOTDEPRCD = 0
TOTDEPREF = 0
FILE_NAME = CURDIR & "\HHT\INL_STR.TXT"
Open FILE_NAME For Input As #1
If EOF(1) Then
    MsgBox "No Bill Value Records for Import!", vbCritical, "DAS Version 4.0"
    Exit Sub
End If

    Do
        Line Input #1, input_data
        TOT_REC = TOT_REC + 1
    Loop While Not EOF(1)
Close #1

Dim fso As New FileSystemObject, txtfile, fil1 As File, ts As TextStream
Set fil1 = fso.GetFile(CURDIR & "\HHT\INL_STR.TXT")
Set TXT_STREAM = fil1.OpenAsTextStream(ForReading)
For i = 0 To TOT_REC - 1
    INL_STR = TXT_STREAM.ReadLine
    ERSI_IDY = Trim(Mid(INL_STR, 106, 10))
    ECAS_AMT = Trim(Mid(INL_STR, 186, 10))
    ECHQ_AMT = Trim(Mid(INL_STR, 196, 10))
    EDEP_RCD = Trim(Mid(INL_STR, 248, 10))
    EDEP_REF = Trim(Mid(INL_STR, 228, 10))
    LINV_IDY = Trim(Mid(INL_STR, 1, 10))
    EORD_IDY = Trim(Mid(INL_STR, 106, 10))
    EINV_DAT = Trim(Mid(INL_STR, 11, 10))
    LCST_IDY = Trim(Mid(INL_STR, 21, 10))
    LGRS_AMT = Trim(Mid(INL_STR, 116, 10))
    LCSD_AMT = Trim(Mid(INL_STR, 136, 10))
    EPKG_CST = Trim(Mid(INL_STR, 146, 10))
    LNET_AMT = Trim(Mid(INL_STR, 156, 10))
    LSMN_IDY = Trim(Mid(INL_STR, 61, 10))
    LROU_IDY = Trim(Mid(INL_STR, 96, 10))
    LSYS_DAT = Trim(Mid(INL_STR, 11, 10))
    LGOD_CAS = Mid(INL_STR, 238, 10)
    LGOD_AMT = Mid(INL_STR, 248, 10)
    LSYS_TIM = Trim(Mid(INL_STR, 259, 10))
    LSRN_AMT = Trim(Mid(INL_STR, 126, 10))
    STKCHECK = Trim(Mid(EORD_IDY, 1, 3))
    ENET_AMT = Trim(Mid(INL_STR, 156, 10))
    EPTY_IDY = Trim(Mid(INL_STR, 21, 10))
    Close
    TOTCASAMT = TOTCASAMT + Val(ECAS_AMT)
    TOTCHQAMT = TOTCHQAMT + Val(ECHQ_AMT)
    TOTDEPRCD = TOTDEPRCD + Val(EDEP_RCD)
    TOTDEPREF = TOTDEPREF + Val(EDEP_REF)
    If ACT_CHK.State = 1 Then ACT_CHK.Close
    ACT_CHK.Open "SELECT * FROM ACT where ACT_IDY = '" & LCST_IDY & "'", con, adOpenKeyset, adLockPessimistic
    If Not ACT_CHK.EOF And LINV_IDY <> "COLLECTION" Then
    '= START OF VAN UPDATE ='
        
    If rs.State = 1 Then rs.Close
    If MOD1 = True Then
    End If
    
    rs.Open "select * from VAN WHERE INV_IDY = '" & EORD_IDY & "' ", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
    rs.MoveFirst
    
    While Not rs.EOF
        If Not rs("VAN_CLS") = "C" Then
            rs("CAS_AMT") = 0
            rs("CHQ_AMT") = 0
            rs("DEP_RCD") = 0
            rs("DEP_REF") = 0
            rs("CAS_AMT") = TOTCASAMT
            rs("CHQ_AMT") = TOTCHQAMT
            rs("DEP_RCD") = TOTDEPRCD
            rs("DEP_REF") = TOTDEPREF
            rs.Update
        End If
        rs.MoveNext
    Wend
    Else
    End If
    '= END OF VAN UPDATE ='
 
    '= START OF INL UPDATE ='
    LTXN_TYP = "SAL"
    If Len(EORD_IDY) > 0 And STKCHECK = "RSI" Then
        ESTK_UPD = "N"
    Else
        ESTK_UPD = "Y"
    End If
    ESTK_TYP = "Salable"
    EFST_FLG = "N"
    ETXN_TYP = "SAL"
    
    
    If MOD1 = True Then
    End If
    If LOGRS.State = 1 Then LOGRS.Close
    LOGRS.Open "SELECT * FROM LOG", con, adOpenKeyset, adLockPessimistic
    LOGRS.MoveLast
    LUSR_IDY = LOGRS(0)
       
    If rs.State = 1 Then rs.Close
    If ARS.State = 1 Then ARS.Close
    If MOD1 = True Then
    End If
    
    str = "Sales"
    ARS.Open "select ACT_IDY from ACT WHERE GRP_CLF = '" & str & "'  ", con, adOpenKeyset, adLockPessimistic
    If ARS.RecordCount > 0 Then
        M_SAL_IDY = ARS(0)
    Else
        M_SAL_IDY = Space(10)
    End If
    EACT_IDY = M_SAL_IDY
    If MOD1 = True Then
    End If
    If rs2.State = 1 Then rs2.Close
    rs2.Open "select * from inl where inv_idy like '" & Trim(LINV_IDY) & "' and txn_typ like 'SAL'", con, adOpenKeyset, adLockPessimistic
    If Not rs2.RecordCount > 0 Then
        rs2.AddNew
        rs2("INV_IDY") = LINV_IDY
        rs2("INV_DAT") = EINV_DAT
        rs2("ORD_IDY") = EORD_IDY
        rs2("ACT_IDY") = M_SAL_IDY
        rs2("CST_IDY") = LCST_IDY
        rs2("GRS_AMT") = LGRS_AMT
        If LCSD_AMT = "" Then
            rs2("CSD_AMT") = "0.00"
        Else
            rs2("CSD_AMT") = LCSD_AMT
        End If
        rs2("NET_AMT") = LNET_AMT
        rs2("SMN_IDY") = LSMN_IDY
        rs2("TXN_TYP") = LTXN_TYP
        rs2("ROU_IDY") = LROU_IDY
        rs2("USR_IDY") = LUSR_IDY
        rs2("SYS_DAT") = EINV_DAT
        rs2("SYS_TIM") = LSYS_TIM
        rs2("SRN_AMT") = LSRN_AMT
        rs2("CSD_PCG") = LCSD_PCG
        rs2("DCH_IDY") = "         "
        rs2("LRN_IDY") = "         "
        rs2("WBL_IDY") = "          "
        rs2("VCH_IDY") = "          "
        rs2("SCH_IDY") = "          "
        rs2("LRN_AMT") = "0.00"
        rs2("SPD_PCG") = "0.00"
        rs2("SPD_AMT") = "0.00"
        rs2("CSD_PCG") = "0.00"
        rs2("EXD_PCG") = "0.00"
        rs2("EXD_AMT") = "0.00"
        rs2("TAX_PCG") = "0.00"
        rs2("TAX_AMT") = "0.00"
        rs2("FRT_AMT") = "0.00"
        rs2("LRN_DAT") = EINV_DAT
        rs2("WBL_DAT") = EINV_DAT
        rs2("PKG_CST") = EPKG_CST
        rs2("TTL_CAS") = "0.00"
        rs2("TXN_RMK") = "         "
        rs2("TXN_TPT") = "         "
        rs2("ROU_OFF") = "0.00"
        rs2("CRT_DAY") = "0"
        rs2("DUE_DAT") = EINV_DAT
        rs2("TXN_RM1") = "          "
        rs2("TXN_RM2") = "         "
        rs2("SCH_CHK") = "Y"
        rs2("SYS_DAT") = EINV_DAT
        rs2.Update
    Else
        'MsgBox "No Records for Update"
    End If
    '= END OF INL UPDATE ='

    '= START OF IRT UPDATE ='
    If Len(EINL_DAT) > 0 Then
       ESTK_UPD = "N"
    Else
       ESTK_UPD = "Y"
    End If
        ESTK_TYP = "Salable"
        EFST_FLG = "N"
        ETXN_TYP = "SAL"
        ESTK_TYP = Trim(ESTK_TYP)
        If IRTRS.State = 1 Then IRTRS.Close
        If MOD1 = True Then
        End If
        IRTRS.Open "SELECT * FROM IRT WHERE INV_IDY LIKE '" & Trim(LINV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        If Not IRTRS.RecordCount > 0 Then
'        IRTRS.MoveLast
        IRTRS.AddNew
        IRTRS("INV_IDY") = LINV_IDY
        IRTRS("INV_DAT") = EINV_DAT
        IRTRS("NET_AMT") = LNET_AMT
        IRTRS("SMN_IDY") = LSMN_IDY
        IRTRS("TXN_TYP") = LTXN_TYP
        IRTRS("PTY_IDY") = EPTY_IDY
        IRTRS("ACT_IDY") = "          "
        IRTRS("DUE_DAT") = EINV_DAT
        IRTRS.Update
    Else
        'MsgBox "No Records for Update"
    End If
    '= END OF IRT UPDATE ='
    
    '= START OF PNL UPDATE ='
    If rs.State = 1 Then rs.Close
    If ARS.State = 1 Then ARS.Close
    If MOD1 = True Then
    End If
    If LOGRS.State = 1 Then LOGRS.Close
    If MOD1 = True Then
    End If
    LOGRS.Open "select * from LOG ", con, adOpenKeyset, adLockPessimistic
    LOGRS.MoveLast
    EUSR_IDY = LOGRS(0)
    If ACTRS.State = 1 Then ACTRS.Close
    If MOD1 = True Then
    End If
    ACTRS.Open "select * from ACT WHERE ACT_IDY = '" & LCST_IDY & "'  ", con, adOpenKeyset, adLockPessimistic
    If ACTRS.RecordCount > 0 Then
    ECST_ADD1 = ACTRS("ACT_AD1")
    ECST_ADD2 = ACTRS("ACT_AD2")
    ELOC_IDY = ACTRS("LOC_IDY")
    EMAR_IDY = ACTRS("MAR_IDY")
    ECAT_IDY = ACTRS("CAT_IDY")
    Else
    ECST_ADD1 = ""
    ECST_ADD2 = ""
    ELOC_IDY = ""
    EMAR_IDY = ""
    ECAT_IDY = ""
    End If
    If LOCRS.State = 1 Then LOCRS.Close
    If MOD1 = True Then
    End If
    LOCRS.Open "select * from LOC WHERE LOC_IDY = '" & ELOC_IDY & "'  ", con, adOpenKeyset, adLockPessimistic
    If MARRS.State = 1 Then MARRS.Close
    If MOD1 = True Then
    End If
    MARRS.Open "select * from MAR WHERE MAR_IDY = '" & EMAR_IDY & "'  ", con, adOpenKeyset, adLockPessimistic
    If CATRS.State = 1 Then CATRS.Close
    If MOD1 = True Then
    End If
    CATRS.Open "select * from CAT WHERE CAT_IDY = '" & ECAT_IDY & "'  ", con, adOpenKeyset, adLockPessimistic
    If SMNRS.State = 1 Then SMNRS.Close
    If MOD1 = True Then
    End If
    SMNRS.Open "select * from SMN WHERE SMN_IDY = '" & LSMN_IDY & "'  ", con, adOpenKeyset, adLockPessimistic
    
    If ARS.State = 1 Then ARS.Close
    If MOD1 = True Then
    End If
   
    str = "Sales"
    ARS.Open "select ACT_IDY from ACT WHERE GRP_CLF = '" & str & "'  ", con, adOpenKeyset, adLockPessimistic
    M_SAL_IDY = ARS(0)
    EACT_IDY = M_SAL_IDY

    If DIRRS1.State = 1 Then DIRRS1.Close
    If MOD1 = True Then
    End If

    If DIRRS.State = 1 Then DIRRS.Close
    If MOD1 = True Then
    End If
    If DIRRS2.State = 1 Then DIRRS2.Close
    If MOD1 = True Then
    End If
    If DIRRS3.State = 1 Then DIRRS3.Close
    If MOD1 = True Then
    End If
    If DIRRS4.State = 1 Then DIRRS4.Close
    If MOD1 = True Then
    End If
    If DIRRS5.State = 1 Then DIRRS5.Close
    If MOD1 = True Then
    End If
    If DIRRS6.State = 1 Then DIRRS6.Close
    If MOD1 = True Then
    End If
    If DIRRS7.State = 1 Then DIRRS7.Close
    If MOD1 = True Then
    End If
    If DIRRS8.State = 1 Then DIRRS8.Close
    If MOD1 = True Then
    End If
    If DIRRS9.State = 1 Then DIRRS9.Close
    If MOD1 = True Then
    End If

    Dim STR1, STR2, STR3, STR4, STR5, STR6, STR7, STR8, STR9, ESYS_RM1, ESYS_RM2, ESYS_RM3, ESYS_RM4, ESYS_GST, ESYS_CST As String
    STR1 = "SYS"
    STR2 = "TT1"
    STR3 = "TT2"
    STR4 = "RM1"
    STR5 = "RM2"
    STR6 = "RM3"
    STR7 = "RM4"
    STR8 = "GST"
    STR9 = "CST"
    DIRRS.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR2 & "'", con, adOpenKeyset, adLockPessimistic
    DIRRS1.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR3 & "'", con, adOpenKeyset, adLockPessimistic
    DIRRS2.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR4 & "'", con, adOpenKeyset, adLockPessimistic
    DIRRS3.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR5 & "'", con, adOpenKeyset, adLockPessimistic
    DIRRS4.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR6 & "'", con, adOpenKeyset, adLockPessimistic
    DIRRS5.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR7 & "'", con, adOpenKeyset, adLockPessimistic
    DIRRS6.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR8 & "'", con, adOpenKeyset, adLockPessimistic
    DIRRS7.Open "select * from DIR WHERE DIR_IDY = '" & STR1 & "' AND DIR_IDX = '" & STR9 & "'", con, adOpenKeyset, adLockPessimistic
    EINS_NME = DIRRS(3)
    EINS_ADR = DIRRS1(3)
    ESYS_RM1 = DIRRS2(3)
    ESYS_RM2 = DIRRS3(3)
    ESYS_RM3 = DIRRS4(3)
    ESYS_RM4 = DIRRS5(3)
    ESYS_GST = DIRRS6(3)
    ESYS_CST = DIRRS7(3)
'    ESYS_RM1 = DIRRS8(3)

    EACT_IDY = M_SAL_IDY

    If rs2.State = 1 Then rs2.Close
    If MOD1 = True Then
    End If
    rs2.Open "select * from INL where INV_IDY = ''", con, adOpenKeyset, adLockPessimistic
    If ACTRS.RecordCount > 0 Then
    ECST_NME = ACTRS("ACT_NME")
    ECST_AD1 = ACTRS("ACT_AD1")
    ECST_AD2 = ACTRS("ACT_AD2")
    ECST_AD3 = ACTRS("ACT_AD3")
    ECST_AD4 = ACTRS("ACT_AD4")
    ECST_CST = ACTRS("ACT_CST")
    ECST_GST = ACTRS("ACT_GST")
    ECST_ZIP = ACTRS("ACT_ZIP")
    ELOC_IDY = ACTRS("LOC_IDY")
    ECAT_IDY = ACTRS("CAT_IDY")
    EMAR_IDY = ACTRS("MAR_IDY")
    Else
    End If
    If SMNRS.RecordCount > o Then
    
    ESMN_NME = SMNRS("SMN_NME")
    Else
    SMN_NME = ""
    End If
    If CATRS.RecordCount > 0 Then
    
    ECAT_NME = CATRS("CAT_NME")
    Else
    ECAT_NME = ""
    End If
    If MARRS.RecordCount > 0 Then
    
    EMAR_NME = MARRS("MAR_NME")
    Else
    EMAR_NME = ""
    End If
    If LOCRS.RecordCount > 0 Then
    
    ELOC_NME = LOCRS("LOC_NME")
    Else
    ELOC_NME = ""
    End If
    
    EACTIDY = LCST_IDY
    
    If CHECKTB("PNL.dbf") = 1 Then
    End If
    If PNLRS.State = 1 Then PNLRS.Close
    If MOD1 = True Then
    End If
    
    PNLRS.Open "SELECT * FROM PNL WHERE INV_IDY LIKE '" & Trim(LINV_IDY) & "' and txn_typ like 'SAL'", con, adOpenKeyset, adLockPessimistic
    If Not PNLRS.RecordCount > 0 Then
'        PNLRS.MoveLast
        PNLRS.AddNew
        PNLRS("INV_IDY") = LINV_IDY
        PNLRS("INV_DAT") = EINV_DAT
        PNLRS("ORD_IDY") = EORD_IDY
        PNLRS("CST_IDY") = LCST_IDY
If Not ECST_NME = "" Then
PNLRS("CST_NME") = ECST_NME
Else
PNLRS("CST_NME") = "         "
End If


If Not ECST_AD1 = "" Then
  PNLRS("CST_AD1") = ECST_AD1
Else
  PNLRS("CST_AD1") = " "
End If

If Not ECST_AD2 = "" Then
  PNLRS("CST_AD2") = ECST_AD2
Else
  PNLRS("CST_AD2") = " "
End If
If Not ECST_AD3 = "" Then
  PNLRS("CST_AD3") = ECST_AD3
Else
  PNLRS("CST_AD3") = " "
End If

If Not ECST_AD4 = "" Then
  PNLRS("CST_AD4") = ECST_AD4
Else
  PNLRS("CST_AD4") = " "
End If
If Not ECST_CST = "" Then
PNLRS("CST_CST") = ECST_CST
Else
PNLRS("CST_CST") = " "
End If
        PNLRS("CST_GST") = ECST_GST
        PNLRS("CST_ZIP") = ECST_ZIP
        PNLRS("SMN_NME") = ESMN_NME
        PNLRS("CAT_NME") = ECAT_NME
        PNLRS("LOC_IDY") = ELOC_IDY
        PNLRS("CAT_IDY") = ECAT_IDY
        PNLRS("GRS_AMT") = LGRS_AMT
        PNLRS("CSD_AMT") = LCSD_AMT
        PNLRS("NET_AMT") = LNET_AMT
        PNLRS("SMN_IDY") = LSMN_IDY
        PNLRS("TXN_TYP") = LTXN_TYP
        PNLRS("ROU_IDY") = LROU_IDY
        PNLRS("USR_NME") = EUSR_IDY
        PNLRS("SYS_DAT") = EINV_DAT
        PNLRS("SYS_DAT") = EINV_DAT
        PNLRS("SYS_TIM") = LSYS_TIM
        PNLRS("SRN_AMT") = LSRN_AMT
        PNLRS("CSD_PCG") = "0.00"
        PNLRS("DCH_IDY") = "         "
        PNLRS("LRN_IDY") = "         "
        PNLRS("LRN_AMT") = "00"
        PNLRS("WBL_IDY") = "          "
        PNLRS("VCH_IDY") = "          "
        PNLRS("SCH_IDY") = "          "
        PNLRS("ACT_IDY") = EACT_IDY
        PNLRS("MAR_NME") = EMAR_NME
        PNLRS("SMN_NME") = ESMN_NME
        PNLRS("LOC_NME") = ELOC_NME
        'PNLRS("CAT_NME") = "              "
        PNLRS("SPD_PCG") = "00"
        PNLRS("SPD_AMT") = "000000000"
        PNLRS("EXD_PCG") = "00"
        PNLRS("EXD_CAL") = " "
        PNLRS("EXD_AMT") = "000000000000"
        PNLRS("TAX_PCG") = "0000"
        PNLRS("TAX_AMT") = "0000"
        PNLRS("FRT_AMT") = "0000"
        PNLRS("LRN_AMT") = "000000000000"
        PNLRS("PKG_CST") = EPKG_CST
        PNLRS("TTL_CAS") = "000000000000"
        PNLRS("ROU_OFF") = "0000"
        PNLRS("CRT_DAY") = "0000"
        PNLRS("SRN_AMT") = "000000000000"
        PNLRS("FRT_TYP") = " "
        PNLRS("DCH_IDY") = "          "
        PNLRS("DCH_DAT") = EINV_DAT
        PNLRS("LRN_IDY") = "       "
        PNLRS("LRN_DAT") = EINV_DAT
        PNLRS("WBL_IDY") = "       "
        PNLRS("WBL_DAT") = EINV_DAT
        PNLRS("TXN_RMK") = "       "
        PNLRS("TXN_TPT") = "       "
        PNLRS("TAX_ACT") = "       "
        PNLRS("ROU_ACT") = "       "
        PNLRS("FRT_ACT") = "       "
        PNLRS("DUE_DAT") = EINV_DAT
        PNLRS("DES_MOD") = " "
        PNLRS("STO_RMK") = "    "
        'PNLRS("ROU_IDY") = "       "
        PNLRS("TXN_RM1") = "       "
        PNLRS("TXN_RM2") = "       "
        PNLRS("VCH_IDY") = "       "
        PNLRS("VCH_DAT") = EINV_DAT
        PNLRS("SCH_CHK") = " "
        PNLRS("SCH_IDY") = "       "
        PNLRS("FRM_DAT") = EINV_DAT
        PNLRS("TO_DAT") = EINV_DAT
        PNLRS("DIS_CHK") = " "
        PNLRS("MAR_IDY") = EMAR_IDY
        PNLRS("INS_NME") = EINS_NME
        PNLRS("INS_ADR") = EINS_ADR
        ESYS_RM1 = Trim(ESYS_RM1)
        If ESYS_RM1 = "" Then
        PNLRS("SYS_RM1") = "    "
        Else
        PNLRS("SYS_RM1") = ESYS_RM1
        End If
        ESYS_RM2 = Trim(Mid(ESYS_RM2, 40))
        ESYS_RM3 = Trim(Mid(ESYS_RM3, 40))
        ESYS_RM4 = Trim(Mid(ESYS_RM4, 40))
        ESYS_GST = Trim(Mid(ESYS_GST, 30))
        ESYS_CST = Trim(Mid(ESYS_CST, 30))
        If ESYS_RM2 = "" Then
        PNLRS("SYS_RM2") = "    "
        Else
        PNLRS("SYS_RM2") = ESYS_RM2
        End If
        If ESYS_RM3 = "" Then
        PNLRS("SYS_RM3") = "    "
        Else
        PNLRS("SYS_RM3") = ESYS_RM3
        End If
        If ESYS_RM4 = "" Then
        PNLRS("SYS_RM4") = "    "
        Else
        PNLRS("SYS_RM4") = ESYS_RM4
        End If
        If ESYS_GST = "" Then
        PNLRS("INS_GST") = "     "
        Else
        PNLRS("INS_GST") = ESYS_GST
        End If
        If ESYS_CST = "" Then
        PNLRS("INS_CST") = "    "
        Else
        PNLRS("INS_CST") = ESYS_CST
        End If
        PNLRS.Update
    Else
        'MsgBox "No Records to Update"
    End If
    '= END OF PNL TABLE =
    
    '= START OF SNV = '
    Dim SHLRS As New ADODB.Recordset
    ERATAPP = "B"
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    If MOD1 = True Then
    End If
    If SHLRS.State = 1 Then SHLRS.Close
    If MOD1 = True Then
    End If

    SHLRS.Open "select * from SHL", con, adOpenKeyset, adLockPessimistic
    If SHLRS.RecordCount > 0 Then
        ESHL_IDY = SHLRS("SHL_IDY")
        EPRD_IDY = SHLRS("PRD_IDY")
    Else
    End If
       
    RS1.Open "select * from SNV where INV_IDY like '" & Trim(LINV_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    If PRS.State = 1 Then PRS.Close
    If MOD1 = True Then
    End If
    Dim PRDRS As New ADODB.Recordset
    If PRDRS.State = 1 Then PRDRS.Close
    If MOD1 = True Then
    End If
    PRDRS.Open "select * from PRD where PRD_IDY like '" & Trim(EPRD_IDY) & "'", con, adOpenKeyset, adLockPessimistic
    Dim INLRS As New ADODB.Recordset
    If INLRS.State = 1 Then INLRS.Close
    If MOD1 = True Then
    End If
    INLRS.Open "select * from INL where INV_IDY like '" & Trim(LINV_IDY) & "' ", con, adOpenKeyset, adLockPessimistic
    MsgBox "LGOD-CAS" & LGOD_CAS
    If Val(LGOD_CAS) > "0" Then
    MsgBox "LGOD-CAS" & LGOD_CAS
    If SHLRS.RecordCount > 0 Then
    If Not RS1.RecordCount > 0 Then
    'RS1.MoveLast
        RS1.AddNew
        RS1("INV_IDY") = EORD_IDY
        RS1("INV_DAT") = EINV_DAT
        If SHLRS.RecordCount > 0 Then
        RS1("PRD_IDY") = SHLRS("PRD_IDY")
        Else
        RS1("PRD_IDY") = EPRD_IDY
        End If
        RS1("TXN_TYP") = "DPI"
        RS1("SHL_IDY") = SHLRS("SHL_IDY")
        RS1("I_SHL_NOS") = "0"
        RS1("I_PCK_NOS") = "0"
        RS1("D_SHL_NOS") = Mid(LGOD_CAS, 1, 3)
        ESUBUNT = Mid("LGOD_CAS", 5, 3)
        ER_PCK_NOS = RS1("D_SHL_NOS") * PRDRS("SUB_UNT")
        ER_PCK_NOS = ER_PCK_NOS + ESUBUNT
        RS1("D_PCK_NOS") = ER_PCK_NOS
        RS1("DP_SHL_NO") = RS1("D_SHL_NO")
        RS1("D_PCK_NOS") = PRDRS("SUB_UNT")
        RS1("DEP_AMT") = LGOD_CAS
        RS1("TXN_DAT") = EINV_DAT
        RS1("TXN_IDY") = EORD_IDY ' VALIDATOION TO BE FOLLOWED
        
        RS1("D_SHL_NOS") = RS1("I_SHL_NOS") - RS1("R_SHL_NOS")
        RS1("D_PCK_NOS") = ETOT_PCS - ER_PCK_NOS
        RS1("TXN_DAT") = EINV_DAT
        RS1("CST_IDY") = INLRS("CST_IDY") ' TO BE CHECK
        RS1("STK_UPD") = "Y"
        RS1("PRD_UPC") = PRDRS("SUB_UNT")
        RS1("PRD_PCK") = PRDRS("PRD_PCK")
    Else
    End If
    Else
    End If
    Else
    End If
    
    '= END OF SNV WITH GOD_CAS =
    End If
    TXT_STREAM.Skip (0)
    Next i
    TXT_STREAM.Close
    '= END OF INL_STR.TXT UPDATE =' End of INL update from Hand Held Generated Text File

    '= START OF INV_STR.TXT UPDATE ='
    TOT_REC = 0
    FILE_NAME = CURDIR & "\HHT\INV_STR.TXT"
    Open FILE_NAME For Input As #1
        Do
            Line Input #1, input_data
            TOT_REC = TOT_REC + 1
        Loop While Not EOF(1)
    Close #1
    Dim INVFSO As New FileSystemObject, txtfile1, fil1INV As File, ts1 As TextStream
    Set fil1INV = INVFSO.GetFile(CURDIR & "\HHT\INV_STR.TXT")
    Set INV_STREAM = fil1INV.OpenAsTextStream(ForReading)
    For i = 0 To TOT_REC - 1
        INV_STR = INV_STREAM.ReadLine
        EINV_IDY = Trim(Mid(INV_STR, 1, 10))
        EBAT_IDY = Trim(Mid(INV_STR, 11, 10))
        EPRD_IDY = Trim(Mid(INV_STR, 21, 10))
        EPRD_QTY = Trim(Mid(INV_STR, 52, 7))
        EPRD_RAT = Trim(Mid(INV_STR, 69, 10))
        ECOL_AMT = Trim(Mid(INV_STR, 79, 10))
        ERSHL_NOS = Mid(INV_STR, 89, 7)
        EPRD_UNT = Trim(Mid(INV_STR, 104, 3))
        EPRD_PCS = Trim(Mid(INV_STR, 107, 3))
        ETOT_PCS = Trim(Mid(INV_STR, 99, 5))
        REFSTK_TYP = Trim(Mid(INV_STR, 51, 1))
        ILDS_IDY = "         "
        ERATAPP = "B"
    Close
    Dim PRD_CHK As New ADODB.Recordset
    If PRD_CHK.State = 1 Then PRD_CHK.Close
    PRD_CHK.Open "SELECT * FROM PRD where PRD_IDY = '" & EPRD_IDY & "'", con, adOpenKeyset, adLockPessimistic
    If Not PRD_CHK.EOF And EINV_IDY <> "COLLECTION" Then
    
   ' Dim ESUBUNT As Integer
    ER_SHL_NOS = Mid(ERSHL_NOS, 1, 3)
    ESUBUNT = Mid(ERSHL_NOS, 5, 3)
   
    If Len(EORD_IDY) > 0 And STKCHECK = "RSI" Then
        ESTK_UPD = "N"
        Else
        ESTK_UPD = "Y"
    End If
        ESTK_TYP = "Salable"
        EFST_FLG = "N"
        ETXN_TYP = "SAL"
    If REFSTK_TYP = "S" Then
       ESTK_TYP = "Salable"
    ElseIf REFSTK_TYP = "T" Then
       ESTK_TYP = "TakeBack"
       EPRD_RAT = "-" & EPRD_RAT
       Else
       ESTK_TYP = "Free"
       EPRD_RAT = "0.00"
    End If
    ESTK_TYP = Trim(ESTK_TYP)
    ERATAPP = "B"

    If CHECKTB("INV.dbf") = 1 Then
    End If

    If RS1.State = 1 Then RS1.Close

    If MOD1 = True Then
    End If
    RS1.Open "select * from INV where INV_IDY like '" & Trim(EINV_IDY) & "' and TXN_TYP like 'SAL' and BAT_IDY like '" & Trim(EBAT_IDY) & "' and STK_TYP like '" & Trim(ESTK_TYP) & "'", con, adOpenKeyset, adLockPessimistic
    If PRS.State = 1 Then PRS.Close
    If MOD1 = True Then
    End If
    PRS.Open "select * from PRD WHERE PRD_IDY like '" & EPRD_IDY & "'", con, adOpenKeyset, adLockPessimistic
    EPRD_UOM = PRS("CON_FAC") / PRS("SUB_UNT")
    EPRD_UOM = EPRD_UOM * ETOT_PCS
    ESUB_UNT = PRS("SUB_UNT")
    TOTREC = RS1.RecordCount
    If Not RS1.RecordCount > 0 Then
        RS1.AddNew
        RS1("INV_IDY") = EINV_IDY
        RS1("BAT_IDY") = EBAT_IDY
        RS1("PRD_QTY") = EPRD_QTY
        RS1("PRD_IDY") = EPRD_IDY
        RS1("PRD_RAT") = EPRD_RAT
        RS1("COL_AMT") = ECOL_AMT
        RS1("BAT_IDY") = EBAT_IDY
        RS1("RAT_APP") = ERATAPP
        RS1("ORD_IDY") = EORD_IDY
        RS1("INV_DAT") = EINV_DAT
        RS1("STK_UPD") = ESTK_UPD
        RS1("STK_TYP") = ESTK_TYP
        RS1("FST_FLG") = EFST_FLG
        RS1("TXN_TYP") = ETXN_TYP
        RS1("PRD_UNT") = EPRD_UNT
        RS1("PRD_PCS") = EPRD_PCS
        RS1("TOT_PCS") = ETOT_PCS
        RS1("PRD_UOM") = EPRD_UOM
        RS1("SUB_UNT") = ESUB_UNT
        RS1.Update
     Else
        'MsgBox "No Records to Update"
    End If
    '= END OF INV UPDATE ='
    '= START OF PNV UPDATE ='

    ERATAPP = "B"
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    If MOD1 = True Then
    End If
    RS1.Open "select * from PNV where INV_IDY like '" & Trim(EINV_IDY) & "' and TXN_TYP like 'SAL' and BAT_IDY like '" & Trim(EBAT_IDY) & "' and STK_TYP like '" & Trim(ESTK_TYP) & "'", con, adOpenKeyset, adLockPessimistic
    If PRS.State = 1 Then PRS.Close
    If MOD1 = True Then
    End If
    PRS.Open "select * from PRD WHERE PRD_IDY = '" & EPRD_IDY & " ' ", con, adOpenKeyset, adLockPessimistic
    If BRS.State = 1 Then BRS.Close
    If MOD1 = True Then
    End If
    If PRS1.State = 1 Then PRS1.Close
    If MOD1 = True Then
    End If
    EBRD_NME = PRS("BRD_NME")
    BRS.Open "select * from BRD WHERE PRD_IDY = '" & EPRD_IDY & " ' ", con, adOpenKeyset, adLockPessimistic
    PRS1.Open "select * from PRD WHERE BRD_NME = '" & EBRD_NME & "'", con, adOpenKeyset, adLockPessimistic
    EPRD_NME = PRS("PRD_NME")
    EGRP_IDY = PRS("GRP_IDY")
    If PGRRS.State = 1 Then PGRRS.Close
    If MOD1 = True Then
    End If
    PGRRS.Open "select * from PGR WHERE GRP_IDY = '" & EGRP_IDY & "'", con, adOpenKeyset, adLockPessimistic
    EGRP_CAT = PRS("BRD_NME")
    EGRP_CLS = PGRRS("GRP_CLS")
    ELST_RAT = "0.00"
    ELST_DIS = "0.00"
    EBAT_NME = BRS("BAT_NME")
    EMRP_RAT = BRS("PRD_MRP")
    EPRD_UOM = PRS("CON_FAC") / PRS("SUB_UNT")
    EPRD_UOM = EPRD_UOM * ETOT_PCS
    ESUB_UNT = PRS("SUB_UNT")
    If Not RS1.RecordCount > 0 Then
        RS1.AddNew
        RS1("INV_IDY") = EINV_IDY
        RS1("BAT_IDY") = EBAT_IDY
        RS1("PRD_QTY") = EPRD_QTY
        RS1("PRD_IDY") = EPRD_IDY
        RS1("PRD_RAT") = EPRD_RAT
        RS1("COL_AMT") = ECOL_AMT
        RS1("BAT_IDY") = EBAT_IDY
        RS1("RAT_APP") = ERATAPP
        If Trim(Mid(EORD_IDY, 1, 3)) = "RSI" Then
        RS1("ORD_IDY") = EORD_IDY
        Else
        RS1("ORD_IDY") = "0"
        End If
        RS1("INV_DAT") = EINV_DAT
        RS1("STK_UPD") = ESTK_UPD
        RS1("STK_TYP") = ESTK_TYP
        RS1("FST_FLG") = EFST_FLG
        RS1("TXN_TYP") = ETXN_TYP
        RS1("PRD_UNT") = EPRD_UNT
        RS1("PRD_PCS") = EPRD_PCS
        RS1("TOT_PCS") = ETOT_PCS
        RS1("PRD_UOM") = EPRD_UOM
        RS1("SUB_UNT") = ESUB_UNT
        RS1("PRD_NME") = EPRD_NME
        RS1("GRP_IDY") = EGRP_IDY
        RS1("GRP_CLS") = EGRP_CLS
        RS1("GRP_CAT") = EGRP_CAT
        RS1("LST_RAT") = ELST_RAT
        RS1("LST_DIS") = ELST_DIS
        RS1("BAT_NME") = EBAT_NME
        RS1("MRP_RAT") = EMRP_RAT
        'RS1("GRP_CAT") = "        "
        RS1("RET_QTY") = "0.000"
        RS1("CPN_RAT") = "0.00"
        RS1("CPN_AMT") = "0.00"
        RS1("ADD_DIS") = "0.00"
        RS1("ADD_AMT") = "0.00"
        RS1("TRD_DIS") = "0.00"
        RS1("DIS_AMT") = "0.00"
        RS1("TRD_TAX") = "0.00"
        RS1("TAX_AMT") = "0.00"
        RS1("OPN_QTY") = "0.00"
        RS1("ISS_QTY") = "0.000"
        RS1("TOT_QTY") = "0.000"
        RS1("SCH_AMT") = "0.00"
        RS1("SCH_IDY") = " "
        RS1.Update
    Else
        MsgBox "No Records to Update"
   End If
   
   '= START OF SNV UPDATE ='
    ERATAPP = "B"
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    If MOD1 = True Then
    End If
    If SHLRS.State = 1 Then SHLRS.Close
    If MOD1 = True Then
    End If
    SHLRS.Open "select * from SHL where PRD_IDY like '" & Trim(EPRD_IDY) & "' ", con, adOpenKeyset, adLockPessimistic
    RS1.Open "select * from SNV where INV_IDY like '" & Trim(EINV_IDY) & "' ", con, adOpenKeyset, adLockPessimistic
    If PRS.State = 1 Then PRS.Close
    If MOD1 = True Then
    End If
    'Dim PRDRS As New ADODB.Recordset
    If PRDRS.State = 1 Then PRDRS.Close
    If MOD1 = True Then
    End If
    PRDRS.Open "select * from PRD where PRD_IDY like '" & Trim(EPRD_IDY) & "' ", con, adOpenKeyset, adLockPessimistic
    'Dim INLRS As New ADODB.Recordset
    If INLRS.State = 1 Then INLRS.Close
    If MOD1 = True Then
    End If
    INLRS.Open "select * from INL where INV_IDY like '" & Trim(EINV_IDY) & "' ", con, adOpenKeyset, adLockPessimistic
    
    If SHLRS.RecordCount > 0 Then
    If Not RS1.RecordCount > 0 Then
        RS1.AddNew
        RS1("INV_IDY") = EINV_IDY
        RS1("INV_DAT") = EINV_DAT
        RS1("PRD_IDY") = EPRD_IDY
        RS1("TXN_TYP") = "SAL"
        RS1("SHL_IDY") = SHLRS("SHL_IDY")
        RS1("I_SHL_NOS") = EPRD_UNT
        RS1("I_PCK_NOS") = ETOT_PCS
        RS1("R_SHL_NOS") = ER_SHL_NOS
        ER_PCK_NOS = ER_SHL_NOS * PRDRS("SUB_UNT")
        ER_PCK_NOS = ER_PCK_NOS + ESUBUNT
        RS1("R_PCK_NOS") = ER_PCK_NOS
        RS1("D_SHL_NOS") = RS1("I_SHL_NOS") - RS1("R_SHL_NOS")
        RS1("D_PCK_NOS") = ETOT_PCS - ER_PCK_NOS
        RS1("TXN_DAT") = EINV_DAT
        RS1("TXN_DAT") = EINV_DAT
        RS1("TXN_IDY") = EORD_IDY
        'RS1("CST_IDY") = ECST_IDY
        RS1("CST_IDY") = "0000000000"
        If "EORDIDY" = "" Then
        RS1("STK_UPD") = "Y"
        Else
        RS1("STK_UPD") = "N"
        End If
        RS1("PRD_UPC") = PRDRS("SUB_UNT")
        RS1("PRD_PCK") = PRDRS("PRD_PCK")
        RS1("CST_IDY") = INLRS("CST_IDY")
        RS1("SMN_IDY") = INLRS("SMN_IDY")
        RS1("ROU_IDY") = INLRS("ROU_IDY")
        RS1("PCK_IDY") = "          "
        RS1("DEP_AMT") = "0"
        RS1("DP_SHL_NOS") = "0"
        RS1("DP_PCK_NOS") = "0"
        RS1("BRK_SHL") = "0"
        RS1("BRK_PCK") = "0"
        RS1.Update
    Else
    End If
    Else
    End If
    End If
    INV_STREAM.Skip (0)
    Next i
    Command10.Enabled = True
    MsgBox "Bill Information from Hand Held updated into DAS Version 4.0 Successfully!", vbInformation, "DAS Version 4.0"
    INV_STREAM.Close
End Sub
Private Sub Command11_Click()
    Command11.Enabled = False
    Command2.SetFocus
    in_str = ""
    TOT_REC = 0
    rec_len = 0
    MSComm1.InBufferCount = 0
    Call settings
    FILE_NAME = "BHD.TXT"
    FILE_NAME = CURDIR & "\HHT" & "\" & FILE_NAME
    If FILE_NAME = "" Then
    Exit Sub
    End If
    Open FILE_NAME For Input As #1
    If EOF(1) Then
       MsgBox "Distributor Header File Information Not Correct or Not Available!", vbCritical, "DAS Version 4.0"
       Exit Sub
    End If
    Do
        Line Input #1, input_data
        TOT_REC = TOT_REC + 1
        rec_len = Len(input_data)
    Loop While Not EOF(1)
    Close #1
    Do
        dumm = DoEvents()
         If MSComm1.PortOpen = True Then 'UPDATE TO CHECK THE PORT STATUS BY PKB
        MSComm1.Output = "|"
        COM_ERR
        delay (1)
        If MSComm1.PortOpen = True Then 'UPDATE TO CHECK THE PORT STATUS BY PKB
        If MSComm1.Input = "@" Then Exit Do
        End If
        Else
        Exit Sub
        End If
    Loop
    For i = 1 To 4
        If Len(TOT_REC) < i Then
            TOT_REC = "0" & TOT_REC
        End If
    Next i
    
    For i = 1 To 4
        MSComm1.Output = Mid(TOT_REC, i, 1)
    Next i
    in_str = ""
    For i = 1 To 3
        If Len(rec_len) < i Then
            rec_len = "0" & rec_len
        End If
    Next i
    For i = 1 To 3
        MSComm1.Output = Mid(rec_len, i, 1)
    Next i
    delay (1)
    f_name = Dir(FILE_NAME)
    For i = 1 To 10
        If i <= Len(f_name) Then
            MSComm1.Output = Mid(f_name, i, 1)
        Else
            MSComm1.Output = " "
        End If
    Next i
    Do
        in_str = MSComm1.Input
    Loop While in_str <> "#"
    Open FILE_NAME For Input As #1
    For i = 1 To TOT_REC
        in_str = ""
        Line Input #1, input_data
        For J = 1 To rec_len
            If J <= Len(input_data) Then
                MSComm1.Output = Mid(input_data, J, 1)
            Else
                MSComm1.Output = " "
            End If
        Next J
        MSComm1.Output = " "
        Do
        in_str = MSComm1.Input
        Loop While in_str <> "#"
    Next i
    Close #1
    MsgBox "Bill Header,Distributor Information Send Successful!", vbInformation, "DAS Version 4.0"
    Command11.Enabled = True
    End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Activate()
LOCKCONTROL
Text1.Locked = False
txt_ad1.Locked = False
txt_ad2.Locked = False
txt_inc.Locked = False
Text2.Locked = True
checkdongle
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

If K > 111 And K < 120 Then If TypeOf M_FRM_HHT.ActiveControl Is TextBox Then M_FRM_HHT.ActiveControl.Text = Trim(FUNKEY(K))
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
        If M_FRM_HHT.ActiveControl.name = "Text1" Then arealist
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And s = 2 Then
        OpenReport CURDIR & "\lists.RP1", "HHT List", 0, "", 6
    ElseIf K = 82 And s = 2 Then
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        SSTab1.TabEnabled(1) = False
        If CHECKTB("HHT.dbf") = 1 Then
            MsgBox UCase(TNAME) & " Missing in " & Me.Caption
            cancel
            Exit Sub
        End If
        If check("HHT", "HHT_IDY", Trim(Text1)) = False Then
            Text1.Enabled = True
            Text1.SetFocus
        Else
            Text1.Enabled = False
            Text2.SetFocus
        End If
    Else
        If M_FRM_HHT.ActiveControl.name = "Text1" Then
            Unload Me
        ElseIf M_FRM_HHT.ActiveControl.name = "SSTab1" Then
        Else
            save
        End If
    End If
End If
End Sub

Private Sub Form_Load()
checkdongle
Command11.Enabled = False
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
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select HHT or Press <Esc> to Quit"
End Sub

Private Sub list1_KeyDown(K As Integer, s As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Then
    If list1.ListItems.count > 0 Then
        Dim lst As ListSubItem
        Set lst = list1.SelectedItem.ListSubItems.Item(1)
        If rs.State = 1 Then rs.Close
        rs.Open "select * from HHT where HHT_IDY like '" & Trim(lst.Text) & "'", PNL, adOpenStatic
        Text1 = Trim(rs("HHT_IDY"))
        Text2 = Trim(rs("HHT_NME"))
        Text1 = Sincrement(Text1)
            
            txt_dis.Text = Trim(rs("DST_NME"))
            txt_ad1.Text = Trim(rs("DST_AD1"))
            txt_ad2.Text = Trim(rs("DST_AD2"))
            txt_tfx.Text = Trim(rs("DST_TFX"))
            txt_inc.Text = Trim(rs("HHT_INC"))
            LOCKCONTROL
            txt_dis.Locked = False
            'txt_inc.Locked = True
            txt_tfx.Locked = False
            txt_ad1.Locked = False
            txt_ad2.Locked = False
            

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
M_MDI.STS_BAR.Panels(1).Text = "Type Hand Held ID or Press <Enter> to get List of Existing Hand Held"
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 27 And Text1 = "" Then Unload Me
End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Hand Held Description"
End Sub

Private Sub TEXT1_KeyDown(KeyCode As Integer, Shift As Integer)
If CHECKTB("HHT.DBF") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
If rs.State = 1 Then rs.Close
        rs.Open " select * from HHT where HHT_IDY like '" & Sincrement(Text1) & "'", con, adOpenStatic
        If rs.RecordCount > 0 Then
            Text1 = Sincrement(Text1)
            Text2 = Trim(rs("HHT_NME"))
            txt_inc = Trim(rs("hht_inc"))
            txt_dis = Trim(rs("dst_nme"))
            txt_ad1 = Trim(rs("dst_ad1"))
            txt_ad2 = Trim(rs("dst_ad2"))
            txt_tfx = Trim(rs("dst_tfx"))
                      
            
            Text1.Enabled = False
            Text2.SetFocus
            Toolbar1.Buttons(12).Enabled = True
            MOD1 = True
        Else
            arealist
        End If
End If
End Sub

Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
   If Len(Text2.Text) > 0 Then
        txt_inc.SetFocus
   Else
        MsgBox ("Hand Held Description Should Not Be Blank!"), vbExclamation, "DAS Version 4.0"
        Text2.SetFocus
   End If
End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    UNLOCKCONTROL
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
    If M_FRM_HHT.ActiveControl.name = "Text1" Then arealist
ElseIf Button.KEY = "d" Then
    delete
ElseIf Button.KEY = "pre" Then
    'SYS_PSW.Tag = "HHT_PRE"
    'SYS_PSW.Show 1
ElseIf Button.KEY = "p" Then
'    Openreport CURDIR & "\lists.RP1", "MIBS List", 0, "", 6
    cancel
ElseIf Button.KEY = "pr" Then
    
ElseIf Button.KEY = "crt" Then
    CREATE
End If
End Sub

Private Sub txt_ad1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
txt_ad2.SetFocus
End If
End Sub

Private Sub txt_ad1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub txt_ad2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
   txt_tfx.SetFocus
End If
End Sub

Private Sub txt_ad2_KeyPress(Ke As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub txt_cur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
txt_dis.SetFocus
End If
End Sub

Private Sub txt_cur_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub


Private Sub txt_dis_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
txt_dis.SetFocus
  If Len(txt_dis.Text) > 0 Then
        txt_ad1.SetFocus
   Else
        MsgBox ("Hand Held Details Should Not Be Blank!"), vbExclamation, "DAS Version 4.0"
        txt_dis.SetFocus
 End If

End If
End Sub

Private Sub txt_dis_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub txt_inc_GotFocus()
txt_inc.MaxLength = 6
End Sub

Private Sub txt_inc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    txt_dis.SetFocus
End If
End Sub
Private Sub txt_inc_KeyPress(K As Integer)
If txt_inc.Locked = False Then
    If K > 47 And K < 58 Or K = 8 Or K = 13 And Val(txt_inc) > 0 Then
    Else
    K = 0
    MsgBox "Only Numeric will allowed"
    txt_inc.SetFocus
    txt_inc.Text = ""
    If Len(txt_inc.Text) > 0 Then
    txt_dis.SetFocus
    End If
    End If
    End If
End Sub

Private Sub txt_inc_LostFocus()
txt_inc.MaxLength = 10
End Sub

Private Sub txt_tfx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        save
    End If
End Sub
Public Sub LOCKCONTROL()
For Each Control In M_FRM_HHT
If TypeOf Control Is TextBox Then
Control.Locked = True
End If
Next
End Sub
Public Sub UNLOCKCONTROL()
For Each Control In M_FRM_HHT
If TypeOf Control Is TextBox Then
Control.Locked = False
End If
Next
End Sub
Public Sub delay(dlytime As Variant)
    Dim delay_timer
    Dim dli
    delay_timer = Timer
    dli = 0
    Do
    dli = dli + 1
       If Timer > (delay_timer + dlytime) Then
          Exit Do
       End If
    Loop
End Sub
    Private Sub settings()
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    If Option1.Value = True Then
        Comport = 1
    Else
        Comport = 2
    End If
    With MSComm1
        .CommPort = Comport
        .InBufferSize = 3072
        .OutBufferSize = 512
        .InBufferCount = 0
        .InputLen = 1
        .settings = "57600,n,8,1"
    End With
    If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
End Sub
'Public Sub COM_ERR()
'   Select Case MSComm1.CommEvent
'      Case comEventBreak   ' A Break was received.
'      MsgBox "A COM Break Detected! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventFrame   ' Framing Error
'      MsgBox "Framing Error! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventOverrun   ' Data Lost.
'      MsgBox "Data Transmission Loss! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventRxOver   ' Receive buffer overflow.
'      MsgBox "COM Port Data Receive Buffer Overflow!, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventRxParity   ' Parity Error.
'      MsgBox "COM Port Parity Error! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventTxFull   ' Transmit buffer full.
'      MsgBox "COM Transmit Buffer Full! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventDCB   ' Unexpected error retrieving DCB
'      MsgBox "Error Retrieving DCB, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvCD   ' Change in the CD line.
'      MsgBox "Change in CD Line, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvCTS   ' Change in the CTS line.
'      MsgBox "Change in CTS Line, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvDSR   ' Change in the DSR line.
'      MsgBox "Change in DSR Line, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvRing   ' Change in the Ring Indicator.
'      MsgBox "Change in Ring Indicator, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvReceive   ' Received RThreshold # of
'      MsgBox "Receive Error, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvSend   ' There are SThreshold number of
'                     ' characters in the transmit
'                     ' buffer.
'      MsgBox "Data is in Transmit Buffer! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvEOF   ' An EOF charater was found in
'                     ' the input stream
'      MsgBox "EOF in Input Stream! Check COM Connections!", vbCritical, "DAS Version 4.0"
'   End Select
'End Sub

Public Sub COM_ERR()
    Dim result As Variant
    Select Case MSComm1.CommEvent
      Case comEventBreak   ' A Break was received.
      result = MsgBox("A COM Break Detected! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventFrame   ' Framing Error
      result = MsgBox("Framing Error! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventOverrun   ' Data Lost.
      result = MsgBox("Data Transmission Loss! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventRxOver   ' Receive buffer overflow.
      result = MsgBox("COM Port Data Receive Buffer Overflow! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventRxParity   ' Parity Error.
      result = MsgBox("COM Port Parity Error! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventTxFull   ' Transmit buffer full.
      result = MsgBox("COM Transmit Buffer Full! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventDCB   ' Unexpected error retrieving DCB
      result = MsgBox("Error Retrieving DCB! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvCD   ' Change in the CD line.
      result = MsgBox("Change in CD Line! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvCTS   ' Change in the CTS line.
      result = MsgBox("Change in CTS Line! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvDSR   ' Change in the DSR line.
      result = MsgBox("Change in DSR Line! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvRing   ' Change in the Ring Indicator.
      result = MsgBox("Change in Ring Indicator! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvReceive   ' Received RThreshold # of
      result = MsgBox("Receive Error! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
      result = MsgBox("Data is in Transmit Buffer! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvEOF   ' An EOF charater was found in
                     ' the input stream
      result = MsgBox("EOF in Input Stream! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
   End Select
   If result = 4 Or result = "" Then
   'Unload Me
   ElseIf result = 2 Then
   Unload Me
   'M_FRM_MIBS.Visible = True
   End If
End Sub

