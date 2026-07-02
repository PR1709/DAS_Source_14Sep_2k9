VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form YER_FRM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Accounting Year Setup"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "YER_FRM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   360
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
            Picture         =   "YER_FRM.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "YER_FRM.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "YER_FRM.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "YER_FRM.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "YER_FRM.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "YER_FRM.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "YER_FRM.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "YER_FRM.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "YER_FRM.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "YER_FRM.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "YER_FRM.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "YER_FRM.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
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
      Height          =   270
      ItemData        =   "YER_FRM.frx":334E
      Left            =   960
      List            =   "YER_FRM.frx":3358
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MS 
      Height          =   3045
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   5371
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   330
      BackColor       =   -2147483633
      BackColorSel    =   14737632
      ForeColorSel    =   8388608
      GridColor       =   8421504
      GridColorFixed  =   4210752
      TextStyle       =   3
      TextStyleFixed  =   1
      FocusRect       =   2
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5040
      _ExtentX        =   8890
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
Attribute VB_Name = "YER_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REC_SET As New ADODB.Recordset

Private Sub Form_Activate()
MS.SetFocus
MS.Col = 0
If REC_SET.State = 1 Then REC_SET.Close
REC_SET.Open "SELECT * FROM YER", CON, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then
    MS.Rows = REC_SET.RecordCount + 1
    For A = 1 To REC_SET.RecordCount
        MS.TextMatrix(A, 0) = Trim(REC_SET("YER_idy"))
        MS.TextMatrix(A, 1) = Format(REC_SET("FRM_DAT"), "dd/MM/yyyy")
        MS.TextMatrix(A, 2) = Format(REC_SET("TO_DAT"), "dd/MM/yyyy")
        MS.TextMatrix(A, 3) = Trim(REC_SET("ACT_PAS"))
        REC_SET.MoveNext
    Next
End If
End Sub

Private Sub Form_KeyDown(K As Integer, Shift As Integer)
    If K = 83 And Shift = 2 Then
        If MsgBox("Confirm Values Save ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbYes Then
            Dim REC_SET As New ADODB.Recordset
            For A = 1 To MS.Rows - 1
                If REC_SET.State = 1 Then REC_SET.Close
                REC_SET.Open "select * from YER where YER_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'", CON, adOpenKeyset, adLockPessimistic
                If REC_SET.RecordCount > 0 Then
                    REC_SET("ACT_PAS") = Trim(MS.TextMatrix(A, 3))
                    REC_SET.MoveNext
                End If
            Next
        Else
            Form_Activate
        End If
        MS.SetFocus
    ElseIf K = 88 And Shift = 2 Then
        Unload Me
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 81 And Shift = 2 Then
        If list1.Visible = True Then list1.Visible = False
        Form_Activate
    End If

End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        MS.SetFocus
    Else
        If MsgBox("Confirm Values Save ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbYes Then
            Dim REC_SET As New ADODB.Recordset
            For A = 1 To MS.Rows - 1
                If REC_SET.State = 1 Then REC_SET.Close
                REC_SET.Open "select * from YER where YER_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'", CON, adOpenKeyset, adLockPessimistic
                If REC_SET.RecordCount > 0 Then
                    REC_SET("ACT_PAS") = Trim(MS.TextMatrix(A, 3))
                    REC_SET.MoveNext
                End If
            Next
        End If
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
MS.ColWidth(0) = 1100 ' year id
MS.ColWidth(1) = 1150  'from date
MS.ColWidth(2) = 1150  'to date
MS.ColWidth(3) = 1650   'act_pas
MS.TextMatrix(0, 0) = "Year"
MS.TextMatrix(0, 1) = "From Date"
MS.TextMatrix(0, 2) = "To Date"
MS.TextMatrix(0, 3) = "Cur'nt Year YN"
End Sub

Private Sub Form_Unload(cancel As Integer)
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub list1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    ms_row = MS.Row
    list1.Visible = False
    If list1.Text = "Y" Then
        For A = 1 To MS.Rows - 1
            MS.TextMatrix(A, 3) = "N"
        Next
        MS.TextMatrix(ms_row, 3) = "Y"
    Else
        MsgBox "Accounting Year Must Be Selected !", vbExclamation, "DAS Version 4.0"
    End If
    MS.SetFocus
End If

End Sub

Private Sub MS_EnterCell()
If MS.Col = 0 Then
M_MDI.STS_BAR.Panels(1).Text = "Accounting Year"
ElseIf MS.Col = 1 Then
M_MDI.STS_BAR.Panels(1).Text = "From Date"
ElseIf MS.Col = 2 Then
M_MDI.STS_BAR.Panels(1).Text = "To Date"
ElseIf MS.Col = 3 Then
M_MDI.STS_BAR.Panels(1).Text = "Type 'Y' or 'N' To Select The Accounting Year"
End If
End Sub

Private Sub MS_GotFocus()
MS.HighLight = flexHighlightAlways
End Sub

Private Sub MS_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 13 Then If MS.Col <> 3 Then MS.Col = MS.Col + 1

If K <> 27 Then
    If K = 89 Or K = 78 Then  'YES OR NO
        If MS.Col = 3 Then
            list1.Visible = True
            list1.Top = MS.Top + MS.CellTop
            list1.Left = MS.CellLeft + MS.Left
            list1.Width = MS.CellWidth
            list1.Height = MS.CellHeight
            MS.AllowUserResizing = flexResizeNone
            list1.Text = MS.Text
            list1.SetFocus
        End If
    End If
End If
End Sub

Private Sub MS_LeaveCell()
If list1.Visible = True Then list1.Visible = False
End Sub

Private Sub MS_Scroll()
If list1.Visible = True Then
    list1.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "s" Then
    If MsgBox("Confirm Values Save ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbYes Then
        Dim REC_SET As New ADODB.Recordset
        For A = 1 To MS.Rows - 1
            If REC_SET.State = 1 Then REC_SET.Close
            REC_SET.Open "select * from YER where YER_idy like '" & Trim(MS.TextMatrix(A, 0)) & "'", CON, adOpenKeyset, adLockPessimistic
            If REC_SET.RecordCount > 0 Then
                REC_SET("ACT_PAS") = Trim(MS.TextMatrix(A, 3))
                REC_SET.MoveNext
            End If
        Next
    Else
        Form_Activate
    End If
        MS.SetFocus
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "c" Then
    If list1.Visible = True Then
        list1.Visible = False
        MS.SetFocus
    End If
    Form_Activate
End If
End Sub
