VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PRD_ANA_FRM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productivity Analysis"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   Icon            =   "PRD_ANA_FRM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Export"
      Height          =   330
      Left            =   7440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   4680
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7005
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MS 
      Height          =   4695
      Left            =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   3
      Cols            =   15
      RowHeightMin    =   330
      BackColor       =   -2147483633
      BackColorSel    =   16744576
      FocusRect       =   2
      GridLines       =   2
      MergeCells      =   2
      AllowUserResizing=   1
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
            Picture         =   "PRD_ANA_FRM.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRD_ANA_FRM.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRD_ANA_FRM.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRD_ANA_FRM.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRD_ANA_FRM.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRD_ANA_FRM.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRD_ANA_FRM.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRD_ANA_FRM.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRD_ANA_FRM.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRD_ANA_FRM.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRD_ANA_FRM.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PRD_ANA_FRM.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
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
            Object.Visible         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "New (Ctrl+N)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "s"
            Object.ToolTipText     =   "Save (Ctrl+S)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "c"
            Object.ToolTipText     =   "Cancel (Ctrl+Q)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Key             =   "ed"
            Object.ToolTipText     =   "View (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Scheme ID"
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
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&File Name"
      Height          =   195
      Left            =   3840
      TabIndex        =   4
      Top             =   555
      Width           =   705
   End
End
Attribute VB_Name = "PRD_ANA_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
CommonDialog1.Filter = "*.txt"
CommonDialog1.DefaultExt = "txt"
CommonDialog1.DialogTitle = "Enter File Name"
CommonDialog1.ShowSave
Text2 = CommonDialog1.FileName
End Sub

Private Sub Form_KeyDown(K As Integer, S As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K > 111 And K < 120 Then If TypeOf PRD_ANA_FRM.ActiveControl Is TextBox Then PRD_ANA_FRM.ActiveControl.Text = Trim(FUNKEY(K))
    If K = 88 And S = 2 Then
        Unload Me
    ElseIf K = 80 And S = 2 Then
        If Me.Tag = "Category" Then
            OpenReport "", CURDIR & "\Retailer Vol. Penetration.RRW", 0, "", 6
        ElseIf Me.Tag = "Segment" Then
            OpenReport "", CURDIR & "\Retailer Vol. Penetration Seg.RRW", 0, "", 6
        ElseIf Me.Tag = "Brand" Then
            OpenReport "", CURDIR & "\Retailer Vol. Penetration Brd.RRW", 0, "", 6
        ElseIf Me.Tag = "Product" Then
            OpenReport "", CURDIR & "\Retailer Vol. Penetration Prd.RRW", 0, "", 6
        End If
    ElseIf K = 112 Then
        M_HELP
    End If
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If Me.ActiveControl.name = "Text2" Then
        Unload Me
    Else
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Command2_Click()
On Error GoTo LBL
If Trim(FPATH) = "" And Text2 = "" Then
    MsgBox "Invalid File Name", vbExclamation, "DAS Version 4.0"
    Text2.SetFocus
    Exit Sub
End If
Open Text2 For Output As #1   'Open file for output.
For r = 1 To MS.Rows - 1
    For C = 0 To MS.Cols - 1
        STR1 = STR1 & Trim(MS.TextMatrix(r, C)) & ";"
    Next
    Write #1, STR1
    STR1 = ""
Next
Close #1
MsgBox "Data Exported to " & Text2, vbExclamation, "DAS Version 4.0"
Exit Sub
LBL:
MsgBox err.Description, vbCritical, "DAS Version 4.0"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Then KeyAscii = 0
If Len(Trim(Text2)) <> 0 Then
    MsgBox "Invalid File Name", vbExclamation, "DAS Version 4.0"
Else
    Command2.SetFocus
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "p" Then
    If Me.Tag = "Category" Then
        OpenReport "", CURDIR & "\Retailer Vol. Penetration.RRW", 0, "", 6
    ElseIf Me.Tag = "Segment" Then
        OpenReport "", CURDIR & "\Retailer Vol. Penetration Seg.RRW", 0, "", 6
    ElseIf Me.Tag = "Brand" Then
        OpenReport "", CURDIR & "\Retailer Vol. Penetration Brd.RRW", 0, "", 6
    ElseIf Me.Tag = "Product" Then
        OpenReport "", CURDIR & "\Retailer Vol. Penetration Prd.RRW", 0, "", 6
    End If
End If
End Sub
