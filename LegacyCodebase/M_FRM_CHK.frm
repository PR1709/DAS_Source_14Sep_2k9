VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form M_FRM_CHK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Optional Modules - DAS"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "M_FRM_CHK.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   363
      Left            =   3960
      TabIndex        =   8
      Top             =   1320
      Width           =   839
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   363
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
      Width           =   839
   End
   Begin VB.Frame Frame1 
      Caption         =   "Check Module"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3375
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   840
         Width           =   180
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   480
         Width           =   180
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Left            =   2520
         TabIndex        =   1
         Top             =   1230
         Width           =   180
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Telecom "
         Height          =   195
         Left            =   1575
         TabIndex        =   6
         Top             =   870
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DAS Mobile"
         Height          =   195
         Left            =   1395
         TabIndex        =   5
         Top             =   1260
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Two Way Packs - Glass A/c"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   510
         Width           =   2025
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5340
      _ExtentX        =   9419
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "New (Ctrl+N)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "ed"
            Object.ToolTipText     =   "View (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Area ID"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "brw"
            Object.ToolTipText     =   "Browse (Ctrl+B)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   4560
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6030
      Top             =   825
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
            Picture         =   "M_FRM_CHK.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CHK.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CHK.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CHK.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CHK.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CHK.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CHK.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CHK.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CHK.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CHK.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CHK.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CHK.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Check Modules as Required"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   2010
   End
End
Attribute VB_Name = "M_FRM_CHK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_KeyPress(K As Integer)
If K = 27 Then
Unload Me
ElseIf K = 13 Then
Check2.SetFocus
End If
End Sub
Private Sub CHECK2_KeyPress(K As Integer)
If K = 27 Then
Unload Me
ElseIf K = 13 Then
Check3.SetFocus
End If
End Sub
Private Sub CHECK3_KeyPress(K As Integer)
If K = 27 Then
Unload Me
ElseIf K = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub Command1_Click()
save
MsgBox "You Should Restart DAS Version 4.0 Software!", vbInformation, "DAS Version 4.0"
End
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
Else
save
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Check1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
Dim RS_DIR As New ADODB.Recordset
If RS_DIR.State = 1 Then RS_DIR.Close
    RS_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'MDL' AND DIR_IDX LIKE 'TWP'", CON, adOpenKeyset, adLockPessimistic
If RS_DIR.RecordCount > 0 Then
Else
    RS_DIR.AddNew
    RS_DIR("DIR_IDY") = "MDL"
    RS_DIR("DIR_IDX") = "TWP"
    RS_DIR("DIR_ALS") = ""
    RS_DIR("DIR_MSG") = ""
    RS_DIR.AddNew
    RS_DIR("DIR_IDY") = "MDL"
    RS_DIR("DIR_IDX") = "TCO"
    RS_DIR("DIR_ALS") = ""
    RS_DIR("DIR_MSG") = ""
    RS_DIR.AddNew
    RS_DIR("DIR_IDY") = "MDL"
    RS_DIR("DIR_IDX") = "MIB"
    RS_DIR("DIR_ALS") = ""
    RS_DIR("DIR_MSG") = ""
    RS_DIR.Update
End If
If RS_DIR.State = 1 Then RS_DIR.Close
RS_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'MDL' AND DIR_IDX LIKE 'TWP'", CON, adOpenKeyset, adLockPessimistic
If RS_DIR.RecordCount > 0 Then
If Trim(RS_DIR("DIR_ALS")) = "Y" Then
   Check1.Value = 1
End If
End If
If RS_DIR.State = 1 Then RS_DIR.Close
RS_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'MDL' AND DIR_IDX LIKE 'TCO'", CON, adOpenKeyset, adLockPessimistic
If RS_DIR.RecordCount > 0 Then
If Trim(RS_DIR("DIR_ALS")) = "Y" Then
   Check2.Value = 1
End If
End If
If RS_DIR.State = 1 Then RS_DIR.Close
RS_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'MDL' AND DIR_IDX LIKE 'MIB'", CON, adOpenKeyset, adLockPessimistic
If RS_DIR.RecordCount > 0 Then
If Trim(RS_DIR("DIR_ALS")) = "Y" Then
   Check3.Value = 1
End If
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    CLEAR
ElseIf Button.KEY = "s" Then
    save
ElseIf Button.KEY = "c" Then
    CLEAR
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "h" Then
  
ElseIf Button.KEY = "ed" Then
    
ElseIf Button.KEY = "d" Then
   CLEAR
ElseIf Button.KEY = "pre" Then
    
ElseIf Button.KEY = "p" Then

ElseIf Button.KEY = "pr" Then
    
ElseIf Button.KEY = "crt" Then
    
End If

End Sub

Public Sub save()
If MsgBox("Confirm Record Save ?", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    CLEAR
    Exit Sub
Else
        Dim RS_DIR As New ADODB.Recordset
        If RS_DIR.State = 1 Then RS_DIR.Close
        RS_DIR.Open "SELECT * FROM DIR WHERE TRIM(DIR_IDY) LIKE 'MDL' AND TRIM(DIR_IDX) LIKE 'TWP'", CON, adOpenKeyset, adLockPessimistic
        If RS_DIR.RecordCount > 0 Then
        If Check1.Value = 1 Then
        RS_DIR("DIR_ALS") = "Y"
        RS_DIR.Update
        End If
        End If

        If RS_DIR.State = 1 Then RS_DIR.Close
        RS_DIR.Open "SELECT * FROM DIR WHERE TRIM(DIR_IDY) LIKE 'MDL' AND TRIM(DIR_IDX) LIKE 'TCO'", CON, adOpenKeyset, adLockPessimistic
        If RS_DIR.RecordCount > 0 Then
        If Check2.Value = 1 Then
        RS_DIR("DIR_ALS") = "Y"
        RS_DIR.Update
        End If
        End If

        If RS_DIR.State = 1 Then RS_DIR.Close
        RS_DIR.Open "SELECT * FROM DIR WHERE TRIM(DIR_IDY) LIKE 'MDL' AND TRIM(DIR_IDX) LIKE 'MIB'", CON, adOpenKeyset, adLockPessimistic
        If RS_DIR.RecordCount > 0 Then
        If Check3.Value = 1 Then
        RS_DIR("DIR_ALS") = "Y"
        RS_DIR.Update
        End If
        End If
End If
CLEAR
Unload Me
End Sub
 
Public Sub CLEAR()
 Check1.Value = False
    Check2.Value = False
    Check3.Value = False
End Sub

