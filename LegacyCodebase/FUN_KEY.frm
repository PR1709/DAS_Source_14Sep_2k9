VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FUN_KEY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Funtion Keys  "
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4635
   Icon            =   "FUN_KEY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   960
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
            Picture         =   "FUN_KEY.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUN_KEY.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUN_KEY.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUN_KEY.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUN_KEY.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUN_KEY.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUN_KEY.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUN_KEY.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUN_KEY.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUN_KEY.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUN_KEY.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FUN_KEY.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1050
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   1852
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
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
      TabCaption(0)   =   "F&2"
      TabPicture(0)   =   "FUN_KEY.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "F&3"
      TabPicture(1)   =   "FUN_KEY.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "F&4"
      TabPicture(2)   =   "FUN_KEY.frx":3386
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "F&5"
      TabPicture(3)   =   "FUN_KEY.frx":33A2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "F&6"
      TabPicture(4)   =   "FUN_KEY.frx":33BE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "F&7"
      TabPicture(5)   =   "FUN_KEY.frx":33DA
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text6"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "F&8"
      TabPicture(6)   =   "FUN_KEY.frx":33F6
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Text7"
      Tab(6).ControlCount=   1
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   720
         MaxLength       =   25
         TabIndex        =   0
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -74280
         MaxLength       =   25
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -74280
         MaxLength       =   25
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -74280
         MaxLength       =   25
         TabIndex        =   5
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -74280
         MaxLength       =   25
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -74280
         MaxLength       =   25
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -74280
         MaxLength       =   25
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s"
            Object.ToolTipText     =   "Save (Ctrl+S)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "c"
            Object.ToolTipText     =   "Cancel (Ctrl+Q)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FUN_KEY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel()
For A = 1 To FUN_KEY.COUNT - 1
    If TypeOf FUN_KEY.Controls(A) Is TextBox Then FUN_KEY.Controls(A).Text = ""
Next
SSTab1.Tab = 0
Text1.SetFocus
End Sub


Private Sub save()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If MsgBox("Confirm Save ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbYes Then
    rs.Open "select * from fun", CON, adOpenKeyset, adLockPessimistic
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs.Update
    End If
    rs("fun_two") = Trim(Text1)
    rs("fun_thr") = Trim(Text2)
    rs("fun_for") = Trim(Text3)
    rs("fun_fiv") = Trim(Text4)
    rs("fun_six") = Trim(Text5)
    rs("fun_sev") = Trim(Text6)
    rs("fun_eig") = Trim(Text7)
    rs.Update
    
    F2 = Trim(rs("fun_two")) & ""
    F3 = Trim(rs("fun_thr")) & ""
    F4 = Trim(rs("fun_for")) & ""
    F5 = Trim(rs("fun_fiv")) & ""
    F6 = Trim(rs("fun_six")) & ""
    F7 = Trim(rs("fun_sev")) & ""
    F8 = Trim(rs("fun_eig")) & ""
End If
SSTab1.Tab = 0
Text1.SetFocus
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
Text1.SetFocus
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from fun", CON, adOpenKeyset, adLockPessimistic
Text1 = Trim(rs("fun_two")) & ""
Text2 = Trim(rs("fun_thr")) & ""
Text3 = Trim(rs("fun_for")) & ""
Text4 = Trim(rs("fun_fiv")) & ""
Text5 = Trim(rs("fun_six")) & ""
Text6 = Trim(rs("fun_sev")) & ""
Text7 = Trim(rs("fun_eig")) & ""
rs.Update
End Sub

Private Sub Form_KeyDown(K As Integer, S As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
ElseIf K = 88 And S = 2 Then
    Unload Me
    Set Form = Nothing
ElseIf K = 83 And S = 2 Then
    save
ElseIf K = 81 And S = 2 Then
    cancel
ElseIf K = 112 Then
    M_HELP
End If
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If FUN_KEY.ActiveControl.name = "Text1" Then
        Unload Me
        Set Form = Nothing
    Else
        save
    End If
End If
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    Text1.SetFocus
ElseIf SSTab1.Tab = 1 Then
    Text2.SetFocus
ElseIf SSTab1.Tab = 2 Then
    Text3.SetFocus
ElseIf SSTab1.Tab = 3 Then
    Text4.SetFocus
ElseIf SSTab1.Tab = 4 Then
    Text5.SetFocus
ElseIf SSTab1.Tab = 5 Then
    Text6.SetFocus
ElseIf SSTab1.Tab = 6 Then
    Text7.SetFocus
End If
End Sub

Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Value for F2"
SendKeys "{home}+{end}"
End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Value for F3"
SendKeys "{home}+{end}"
End Sub

Private Sub Text3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Value for F4"
SendKeys "{home}+{end}"
End Sub

Private Sub Text4_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Value for F5"
SendKeys "{home}+{end}"
End Sub

Private Sub Text5_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Value for F6"
SendKeys "{home}+{end}"
End Sub

Private Sub Text6_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Value for F7"
SendKeys "{home}+{end}"
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K <= 123 Then K = K - 32
If K = 13 Then
    SSTab1.Tab = 1
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K <= 123 Then K = K - 32
If K = 13 Then
    SSTab1.Tab = 2
    Text3.SetFocus
End If
End Sub

Private Sub text3_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K <= 123 Then K = K - 32
If K = 13 Then
    SSTab1.Tab = 3
    Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K <= 123 Then K = K - 32
If K = 13 Then
    SSTab1.Tab = 4
    Text5.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K <= 123 Then K = K - 32
If K = 13 Then
    SSTab1.Tab = 5
    Text6.SetFocus
End If
End Sub

Private Sub Text7_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Value for F8"
End Sub

Private Sub Text6_KeyPress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K <= 123 Then K = K - 32
If K = 13 Then
    SSTab1.Tab = 6
    Text7.SetFocus
End If
End Sub

Private Sub Text7_keypress(K As Integer)
If K = 39 Or K = 41 Then K = 0
If K > 96 And K <= 123 Then K = K - 32
If K = 13 Then
    SSTab1.Tab = 0
    Text1.SetFocus
End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "e" Then
    Unload Me
    Set Form = Nothing
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "s" Then
    save
ElseIf Button.KEY = "c" Then
    cancel
End If
End Sub



