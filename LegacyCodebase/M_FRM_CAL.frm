VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_CAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Calender"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "M_FRM_CAL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8880
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbPeriod 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "M_FRM_CAL.frx":0442
      Left            =   1680
      List            =   "M_FRM_CAL.frx":0455
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Generate Default Calendar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DATE2 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   74514433
      CurrentDate     =   37753
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   74514435
      CurrentDate     =   37728
   End
   Begin MSFlexGridLib.MSFlexGrid ms 
      Height          =   4095
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7223
      _Version        =   393216
      Rows            =   53
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   330
      BackColor       =   -2147483633
      BackColorSel    =   12632256
      FocusRect       =   0
      GridLines       =   2
      ScrollBars      =   2
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":0485
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":09C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":0D91
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":10E5
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":171D
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":1A71
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":1ED1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":21C5
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":25D1
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":26DD
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":2A31
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":2E45
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_CAL.frx":3391
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
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
            Object.Visible         =   0   'False
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "WSR"
                  Text            =   "Weekly Sales Report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EWSR"
                  Text            =   "Extended Week Report"
               EndProperty
            EndProperty
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
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pre"
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
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   74514433
      CurrentDate     =   37753
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Period Frequency"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   855
      Width           =   1245
   End
End
Attribute VB_Name = "M_FRM_CAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cls1 As New Class1
Dim mn As Integer, no As Integer, nod As Integer, dacount As Integer
Dim yr As Integer
Private Sub Command1_Click()
Cls1.SaveCalender
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub cmbPeriod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(cmbPeriod.Text)) > 0 Then
   
        MS.Enabled = True
        If cmbPeriod.ListIndex = 0 Then
            cmbPeriod.ItemData(0) = 6
            RN = 52
            MS.Rows = 53
            scw
        ElseIf cmbPeriod.ListIndex = 1 Then
            cmbPeriod.ItemData(1) = 14
            RN = 26
            MS.Rows = 27
            scw
        ElseIf cmbPeriod.ListIndex = 2 Then
            mn = 1
            RN = 12
            MS.Rows = 13
            scw
        ElseIf cmbPeriod.ListIndex = 3 Then
            mn = 3
            RN = 4
            MS.Rows = 5
            MS.ColWidth(0) = 800
            MS.ColWidth(1) = 2000
            MS.ColWidth(2) = 1490
            MS.ColWidth(3) = 2000
            MS.ColWidth(4) = 1490
            MS.ColWidth(5) = 0
            
            Dim r As Integer
            Dim C As Integer
            For r = 1 To MS.Rows - 1
                For C = 0 To MS.Cols - 1
                MS.TextMatrix(r, C) = ""
                Next
            Next
            
        ElseIf cmbPeriod.ListIndex = 4 Then
            mn = 6
            RN = 2
            MS.Rows = 3
            MS.ColWidth(0) = 800
            MS.ColWidth(1) = 2000
            MS.ColWidth(2) = 1490
            MS.ColWidth(3) = 2000
            MS.ColWidth(4) = 1490
            MS.ColWidth(5) = 0
            
            For r = 1 To MS.Rows - 1
                For C = 0 To MS.Cols - 1
                MS.TextMatrix(r, C) = ""
                Next
            Next
            
        End If
        MS.Row = 1
        MS.Col = 1
        MS.TextMatrix(1, 0) = 1
        MS.SetFocus
    Else
        MsgBox "Select Calender Period!", vbExclamation, "DAS Version 4.0"
        cmbPeriod.SetFocus
        MS.Enabled = False
    End If
    'ms.Rows = 53
    If Len(Trim(cmbPeriod.Text)) > 0 Then
        Command3.Enabled = True
    End If
End If
'If Len(Trim(cmbPeriod.Text)) > 0 Then
'    Command3.Enabled = True
'End If
End Sub

Private Sub Command3_Click()
If cmbPeriod.ListIndex = 0 Or cmbPeriod.ListIndex = 1 Then
If MS.Rows > 1 Then
    If Len(Trim(MS.TextMatrix(1, 2))) > 0 And Len(Trim(MS.TextMatrix(1, 4))) > 0 Then
    GENERATECALENDER
    End If
Else
    MS.Rows = MS.Rows + 1
    GENERATECALENDER
End If
Else
    GENERATECALENDER
End If
Command3.Enabled = False
End Sub

Private Sub date1_GotFocus()
DATE1.Value = Date  ' CHANGED BY PKB
End Sub
Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
DATE1.Refresh
If KeyCode = 13 Then
    DT = MS.Text
    MS.Text = Format(DATE1.Value, "dd/mmm/yyyy")
    DATE2.Value = Format(MS.Text, "dd/mmm/yyyy")
    DATE1.Refresh
    If MS.Row = 1 And MS.Col = 1 Then
            DATE1.Visible = False
            MS.Text = DATE1.Value
            MS.TextMatrix(MS.Row, 2) = Format(DATE1, "dddd")
            MS.TextMatrix(MS.Row, 1) = Format(DATE1, "dd/mmm/yyyy")
            DTPicker1.Value = Format(MS.TextMatrix(MS.Row, 1), "DD/MMM/YYYY")
            MS.Col = 3
            MS.SetFocus
    ElseIf MS.Col = 3 Then
        If DateDiff("d", Format(DATE2.Value, "dd/mmm/yyyy"), Format(DATE1.Value, "dd/mmm/yyyy")) < 0 Then
            MsgBox "Invalid Date", vbExclamation, "DAS Version 4.0"
            MS.Text = DT
            DATE1.Visible = False
            MS.SetFocus
            Exit Sub
        Else
            MS.SetFocus
            MS.Text = Format(DATE1.Value, "dd/mmm/yyyy")
            MS.TextMatrix(MS.Row, 4) = Format(DATE1.Value, "dddd")
            DATE1.Refresh
            If DateDiff("d", Format(MS.TextMatrix(MS.Row, 1), "dd/mmm/yyyy"), Format(DATE1, "dd/mmm/yyyy")) < 0 Then
                MS.Text = DT
                MsgBox "Invalid Date", vbExclamation, "DAS Version 4.0"
                DATE1.Visible = False
                MS.SetFocus
                Exit Sub
            End If
            If IsDate(MS.Text) = True Then
                If MS.Row <> MS.Rows - 1 Then
                    MS.TextMatrix(MS.Row + 1, 1) = Format(DateAdd("D", 1, DATE1), "dd/mmm/yyyy")
                    MS.TextMatrix(MS.Row + 1, 2) = Format(DateAdd("D", 1, DATE1), "dddd")
                    DATE1.Visible = False
                    MS.SetFocus
                    Exit Sub
                End If
            Else
                If MS.Row <> MS.Rows - 1 Then
                    MS.TextMatrix(MS.Row + 1, 1) = Format(DateAdd("D", 1, DATE1), "dd/mmm/yyyy")
                    MS.TextMatrix(MS.Row + 1, 2) = Format(DateAdd("D", 1, DATE1), "dddd")
                    DATE1.Visible = False
                    MS.SetFocus
                    Exit Sub
                End If
            End If
            DATE1.Visible = False
            MS.SetFocus
        End If
    End If
End If
End Sub
Private Sub Form_Activate()
cmbPeriod.Enabled = True
cmbPeriod.SetFocus
End Sub

Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If

If K = 83 And s = 2 Then
    If MsgBox("Confirm Save FMCG Calender ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
        Cls1.SaveCalender
    End If
ElseIf K = 81 And s = 2 Then

    MS.Rows = 1
    'ms.Rows = 53
'    For A = 1 To ms.Rows - 1
'        ms.Col = 0
'        ms.TextMatrix(A, 0) = A
'        ms.Row = A
'        ms.CellFontBold = True
'    Next
    cmbPeriod.Enabled = True
    ' MODIFIED BY SVS
'    Command3.Enabled = True
    DATE1.Visible = False
    DATE2.Visible = False
    Text1.Visible = False
    DTPicker1.Visible = False
    cmbPeriod.SetFocus
ElseIf K = 88 And s = 2 Then
    Unload Me
ElseIf K = 68 And s = 2 Then
    If CHECKTB("wek.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        Exit Sub
    End If
    If MsgBox("Are you sure to Delete FMCG Calender ?", vbExclamation + vbYesNo, "DAS Version 4.0") = vbYes Then
        con.Execute "delete from wek"
        MS.Rows = 1
        MS.Rows = 53
        For A = 1 To MS.Rows - 1
            MS.TextMatrix(A, 0) = A
            MS.Row = A
            MS.Col = 0
            MS.CellFontBold = True
        Next
        ' MODIFIED BY SVS
        'Command3.Enabled = True
    End If
ElseIf K = 112 Then
    M_HELP
ElseIf K = 80 And s = 2 Then
    OpenReport CURDIR & "\lists.RP1", "Calender List", 0, "", 6
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If MsgBox("Confirm Save FMCG Calender", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
        Cls1.SaveCalender
    End If
    Unload Me
End If
End Sub

Private Sub Form_Load()
If CHECKTB("wek.dbf") = 1 Then
        MsgBox UCase(TNAME) & " Missing in " & Me.Caption
        Exit Sub
End If
MS.TextMatrix(0, 0) = "Period"
SCl MS, 0
MS.TextMatrix(0, 1) = "Period Start"
SCl MS, 1
MS.TextMatrix(0, 2) = "Day"
MS.TextMatrix(0, 3) = "Period End"
SCl MS, 3
MS.TextMatrix(0, 4) = "Day"
MS.TextMatrix(0, 5) = "Month"
SCl MS, 5
MS.ColWidth(0) = 600
MS.ColWidth(1) = 1700
MS.ColAlignment(1) = 1
MS.ColWidth(2) = 1325
MS.ColWidth(3) = 1700
MS.ColAlignment(3) = 1
MS.ColWidth(4) = 1325
MS.ColWidth(5) = 1130
MS.ColAlignment(5) = 7
MS.TextMatrix(1, 0) = 1
'For A = 1 To ms.Rows - 1
'ms.TextMatrix(A, 0) = A
'ms.Row = A
'ms.Col = 0
'ms.CellFontBold = True
'Next
Dim rs As New ADODB.Recordset
rs.Open "select * from wek", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    Command3.Enabled = False
    MS.Rows = rs.RecordCount + 1
    For A = 1 To rs.RecordCount
        MS.TextMatrix(A, 0) = rs(0)
        MS.TextMatrix(A, 1) = Format(rs(1), "dd/mmm/yyyy")
        MS.TextMatrix(A, 2) = Format(rs(1), "dddd")
        MS.TextMatrix(A, 3) = Format(rs(2), "dd/mmm/yyyy")
        MS.TextMatrix(A, 4) = Format(rs(2), "dddd")
        MS.TextMatrix(A, 5) = rs(3)
        rs.MoveNext
    Next
Else
    cmbPeriod.Enabled = True
End If
If rs.RecordCount = 52 Then
    cmbPeriod.ListIndex = 0
ElseIf rs.RecordCount = 26 Then
    cmbPeriod.ListIndex = 1
ElseIf rs.RecordCount = 12 Then
    cmbPeriod.ListIndex = 2
ElseIf rs.RecordCount = 4 Then
    cmbPeriod.ListIndex = 3
ElseIf rs.RecordCount = 2 Then
    cmbPeriod.ListIndex = 4
End If
MS.Col = 1
MS.Row = 1
End Sub

Private Sub MS_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
If K <> 27 Then
    If Len(Trim(cmbPeriod.Text)) > 0 Then
    If (MS.Row = 1 And MS.Col = 1) Or (MS.Col = 3) Then
        If MS.Col = 3 Then If IsDate(Format(MS.TextMatrix(MS.Row, 1))) = False Then Exit Sub
        If IsDate(MS.Text) = True Then DATE1 = MS.Text
        DATE1.Visible = True
        DATE1.Top = MS.Top + MS.CellTop
        DATE1.Left = MS.CellLeft + MS.Left
        DATE1.Width = MS.CellWidth
        DATE1.Height = MS.CellHeight
        DATE1.Visible = True
        DATE1.SetFocus
    ElseIf MS.Col = 5 Or MS.Col = 0 Then
        If K > 47 And K < 58 Or K = 8 Or K = 46 Then
        Else
            K = 0
        End If
        If K <> 0 Then
            Text1.Visible = True
            Text1.Top = MS.Top + MS.CellTop
            Text1.Left = MS.CellLeft + MS.Left
            Text1.Width = MS.CellWidth
            Text1.Height = MS.CellHeight
            Text1 = Chr(K)
            Text1.SelStart = Len(Text1)
            Text1.Visible = True
            Text1.SetFocus
        End If
    ElseIf MS.Col = 0 And MS.Row = 1 Then
        If K > 47 And K < 58 Or K = 8 Or K = 46 Then
        Else
            K = 0
        End If
        If K <> 0 Then
            Text1.Visible = True
            Text1.Top = MS.Top + MS.CellTop
            Text1.Left = MS.CellLeft + MS.Left
            Text1.Width = MS.CellWidth
            Text1.Height = MS.CellHeight
            Text1 = Chr(K)
            Text1.SelStart = Len(Text1)
            Text1.Visible = True
            Text1.SetFocus
        End If
    End If
    End If
End If
End Sub

Private Sub MS_LeaveCell()
DATE1.Visible = False
End Sub

Private Sub MS_Scroll()
If DATE1.Visible = True Then
    DATE1.Visible = False
    MS.SetFocus
ElseIf DATE2.Visible = True Then
    DATE2.Visible = False
    MS.SetFocus
ElseIf Text1.Visible = True Then
    Text1 = ""
    Text1.Visible = False
    MS.SetFocus
End If
End Sub

Private Sub TEXT1_KeyDown(K As Integer, s As Integer)
If K = 13 Then
    If MS.Col = 0 Then
        If Val(Text1) < RN And Val(Text1) > 0 Then
            MS.Text = Text1
            Text1.Visible = False
            MS.Col = 1
            i = MS.Row + 1
            Do While 1
                If Len(Trim(MS.TextMatrix(i, 1))) > 0 Then
                    MS.TextMatrix(i, 0) = Val(MS.TextMatrix(i - 1, 0)) + 1
                Else
                    Exit Do
                End If
            i = i + 1
            If i - 1 = MS.Rows - 1 Then Exit Do
            Loop
            MS.SetFocus
        Else
            MsgBox "Invalid Week Number", vbExclamation
            Text1.Visible = False
        End If
    Else
        If Val(Text1) < 13 And Val(Text1) > 0 Then
            MS.Text = Text1
            Text1.Visible = False
            If MS.Rows - 1 <> MS.Row Then MS.Row = MS.Row + 1
        Else
            MsgBox "Invalid Month Number", vbExclamation
            Text1.Visible = False
        End If
        MS.SetFocus
    End If
End If
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
    K = 0
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "s" Then
    If MsgBox("Confirm Save FMCG Calender ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
        Cls1.SaveCalender
    End If
ElseIf Button.KEY = "d" Then
    If MsgBox("Are you sure to Delete FMCG Calender", vbExclamation + vbYesNo, "DAS Version 4.0") = vbYes Then
        con.Execute "delete from wek"
        MS.Rows = 1
        MS.Rows = 53
        For A = 1 To MS.Rows - 1
            MS.TextMatrix(A, 0) = A
            MS.Row = A
            MS.Col = 0
            MS.CellFontBold = True
        Next
        ' MODIFIED BY SVS
'        Command3.Enabled = True
    End If
ElseIf Button.KEY = "c" Then
    MS.Rows = 1
    cmbPeriod.Enabled = True
    DATE1.Visible = False
    DATE2.Visible = False
    Text1.Visible = False
    DTPicker1.Visible = False
    cmbPeriod.SetFocus
'''    MS.Rows = 1
'''    MS.Rows = RN
'''    For A = 1 To MS.Rows - 1
'''        MS.TextMatrix(A, 0) = A
'''        MS.Row = A
'''        MS.Col = 0
'''        MS.CellFontBold = True
'''    Next
'''    ' MODIFIED BY SVS
''''    Command3.Enabled = True
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "p" Then
    OpenReport CURDIR & "\lists.RP1", "Calender List", 0, "", 6
ElseIf Button.KEY = "h" Then
    M_HELP
End If
End Sub
'Public Sub GENERATECALENDER()
'With ms
'For A = 2 To 52
'    If Val(.TextMatrix(A - 1, 0)) = 52 Then
'        .TextMatrix(A, 0) = 1
'    Else
'        .TextMatrix(A, 0) = Val(.TextMatrix(A - 1, 0) + 1)
'    End If
'    .TextMatrix(A, 1) = Format(DateAdd("D", 1, .TextMatrix(A - 1, 3)), "DD/MMM/YYYY")
'    .TextMatrix(A, 2) = Format(DateAdd("D", 1, .TextMatrix(A - 1, 3)), "DDDD")
'    .TextMatrix(A, 3) = Format(DateAdd("D", 6, .TextMatrix(A, 1)), "DD/MMM/YYYY")
'    .TextMatrix(A, 4) = Format(DateAdd("D", 6, .TextMatrix(A, 1)), "DDDD")
'Next
'For A = 1 To .Rows - 1
'    .TextMatrix(A, 5) = Val(Format(.TextMatrix(A, 1), "MM"))
'Next
'End With
'End Sub
Public Sub GENERATECALENDER()
If cmbPeriod.ListIndex = 0 Or cmbPeriod.ListIndex = 1 Then

    With MS
    .TextMatrix(1, 0) = 1
    For A = 2 To RN
        .TextMatrix(A, 0) = A
        If Val(.TextMatrix(A - 1, 0)) = RN Then
            .TextMatrix(A, 0) = 1
        Else
            .TextMatrix(A, 0) = Val(.TextMatrix(A - 1, 0) + 1)
        End If
        .TextMatrix(A, 1) = Format(DateAdd("D", 1, .TextMatrix(A - 1, 3)), "DD/MMM/YYYY")
        .TextMatrix(A, 2) = Format(DateAdd("D", 1, .TextMatrix(A - 1, 3)), "DDDD")
        .TextMatrix(A, 3) = Format(DateAdd("D", cmbPeriod.ItemData(cmbPeriod.ListIndex), .TextMatrix(A, 1)), "DD/MMM/YYYY")
        .TextMatrix(A, 4) = Format(DateAdd("D", cmbPeriod.ItemData(cmbPeriod.ListIndex), .TextMatrix(A, 1)), "DDDD")
    Next
    For A = 1 To .Rows - 1
        .TextMatrix(A, 5) = Val(Format(.TextMatrix(A, 1), "MM"))
    Next
    End With
ElseIf cmbPeriod.ListIndex = 2 Then

    With MS

        .TextMatrix(1, 0) = 1
        nodays Month(DTPicker1.Value), Year(DTPicker1.Value)
        DTPicker1.Day = dacount
        .TextMatrix(1, 3) = Format(DTPicker1.Value, "DD/MMM/YYYY") ' Format(DateAdd("D", 29, .TextMatrix(1, 1)), "DD/MMM/YYYY")
        .TextMatrix(1, 4) = Format(DTPicker1.Value, "DDDD") 'Format(DateAdd("D", 29, .TextMatrix(1, 1)), "DDDD")
        DTPicker1.Value = Format(.TextMatrix(1, 3), "DD/MMM/YYYY")
        For A = 2 To RN
            .TextMatrix(A, 0) = A
            DTPicker1.Day = 1
            If DTPicker1.Month = 12 Then
                DTPicker1.Month = 1
            Else
                DTPicker1.Month = DTPicker1.Month + 1
            End If
            DTPicker1.Refresh
            .TextMatrix(A, 1) = Format(DTPicker1.Value, "DD/MMM/YYYY")
            .TextMatrix(A, 2) = Format(DTPicker1.Value, "DDDD")
            nodays Format(DTPicker1.Value, "MM"), Format(DTPicker1.Value, "YYYY")
            DTPicker1.Day = dacount
            .TextMatrix(A, 3) = Format(DTPicker1.Value, "DD/MMM/YYYY")
            .TextMatrix(A, 4) = Format(DTPicker1.Value, "DDDD")
        Next
    For A = 1 To RN
        .TextMatrix(A, 5) = Val(Format(.TextMatrix(A, 1), "MM"))
    Next
    End With
ElseIf cmbPeriod.ListIndex = 3 Then
    loadqtr
ElseIf cmbPeriod.ListIndex = 4 Then
    loadhalf
End If
If MS.Enabled = True Then MS.SetFocus
MS.Row = 1
MS.Col = 3
End Sub

Public Sub nodays(M As Integer, y As Integer)
If M = 1 Or M = 3 Or M = 5 Or M = 7 Or M = 8 Or M = 10 Or M = 12 Then
    dacount = 31
ElseIf M = 4 Or M = 6 Or M = 9 Or M = 11 Then
    dacount = 30
ElseIf M = 2 Then
    If y Mod 4 = 0 Then
        dacount = 29
    Else
        dacount = 28
    End If
End If
End Sub
Private Sub loadqtr()
With MS
    .TextMatrix(1, 0) = 1
    If DTPicker1.Month >= 11 Then
        DTPicker1.Month = 2 - (12 - Val(DTPicker1.Month))
        DTPicker1.Year = DTPicker1.Year + 1
    Else
        DTPicker1.Month = DTPicker1.Month + 2
    End If
    nodays Val(Format(DTPicker1.Value, "MM")), Val(Format(DTPicker1.Value, "YYYY"))
    DTPicker1.Day = dacount
    .TextMatrix(1, 3) = Format(DTPicker1.Value, "DD/MMM/YYYY") ' Format(DateAdd("D", 29, .TextMatrix(1, 1)), "DD/MMM/YYYY")
    .TextMatrix(1, 4) = Format(DTPicker1.Value, "DDDD") 'Format(DateAdd("D", 29, .TextMatrix(1, 1)), "DDDD")
    DTPicker1.Value = Format(.TextMatrix(1, 3), "DD/MMM/YYYY")
    If DTPicker1.Month = 3 Then
    Else
        For A = 2 To RN
            .TextMatrix(A, 0) = A
            DTPicker1.Day = 1
            If DTPicker1.Month = 12 Then
                DTPicker1.Month = 1
                DTPicker1.Year = DTPicker1.Year + 1
            Else
                DTPicker1.Month = DTPicker1.Month + 1
            End If
            DTPicker1.Refresh
                .TextMatrix(A, 1) = Format(DTPicker1.Value, "DD/MMM/YYYY")
                .TextMatrix(A, 2) = Format(DTPicker1.Value, "DDDD")
                If DTPicker1.Month >= 11 Then
                    DTPicker1.Month = 2 - (12 - Val(DTPicker1.Month))
                    DTPicker1.Year = DTPicker1.Year + 1
                Else
                    DTPicker1.Month = DTPicker1.Month + 2
                End If
                nodays Format(DTPicker1.Value, "MM"), Format(DTPicker1.Value, "YYYY")
                DTPicker1.Day = dacount
                If (DTPicker1.Month >= 3 And DTPicker1.Year = 2004) Then
                    DTPicker1.Month = 3
                    DTPicker1.Day = 31
                    .TextMatrix(A, 3) = Format(DTPicker1.Value, "DD/MMM/YYYY")
                    .TextMatrix(A, 4) = Format(DTPicker1.Value, "DDDD")
                Else
                    .TextMatrix(A, 3) = Format(DTPicker1.Value, "DD/MMM/YYYY")
                    .TextMatrix(A, 4) = Format(DTPicker1.Value, "DDDD")
                End If
                If DTPicker1.Month = 3 Then
                    Exit For
                End If
        Next
    End If
    For A = 1 To RN
        .TextMatrix(A, 5) = Val(Format(.TextMatrix(A, 1), "MM"))
    Next
    End With
End Sub
Private Sub loadhalf()
With MS
    .TextMatrix(1, 0) = 1
    If DTPicker1.Month >= 1 And DTPicker1.Year = 2004 Then
            DTPicker1.Month = 3
            DTPicker1.Day = 31
    Else
        If DTPicker1.Month >= 8 Then
                DTPicker1.Month = 5 - (12 - Val(DTPicker1.Month))
                DTPicker1.Year = DTPicker1.Year + 1
                DTPicker1.Year = 2004
                If DTPicker1.Month >= 4 And DTPicker1.Year = 2004 Then
                    DTPicker1.Month = 3
                    DTPicker1.Day = 31
                Else
                'DTPicker1.Year = DTPicker1.Year + 1
                End If
        Else
            DTPicker1.Month = DTPicker1.Month + 5
        End If
    End If
    nodays Val(Format(DTPicker1.Value, "MM")), Val(Format(DTPicker1.Value, "YYYY"))
    DTPicker1.Day = dacount
    .TextMatrix(1, 3) = Format(DTPicker1.Value, "DD/MMM/YYYY") ' Format(DateAdd("D", 29, .TextMatrix(1, 1)), "DD/MMM/YYYY")
    .TextMatrix(1, 4) = Format(DTPicker1.Value, "DDDD") 'Format(DateAdd("D", 29, .TextMatrix(1, 1)), "DDDD")
    DTPicker1.Value = Format(.TextMatrix(1, 3), "DD/MMM/YYYY")
    .TextMatrix(1, 5) = Val(Format(.TextMatrix(1, 1), "MM"))
    If DTPicker1.Month = 3 Then
    Else
    For A = 2 To RN
        .TextMatrix(A, 0) = A
        DTPicker1.Day = 1
        If DTPicker1.Month = 12 Then
            DTPicker1.Month = 1
            DTPicker1.Year = DTPicker1.Year + 1
        Else
            DTPicker1.Month = DTPicker1.Month + 1
        End If
        DTPicker1.Refresh
        .TextMatrix(A, 1) = Format(DTPicker1.Value, "DD/MMM/YYYY")
        .TextMatrix(A, 2) = Format(DTPicker1.Value, "DDDD")
        If DTPicker1.Month >= 1 And DTPicker1.Year = 2004 Then
            DTPicker1.Month = 3
            DTPicker1.Day = 31
        Else
            If DTPicker1.Month >= 8 Then
                DTPicker1.Month = 5 - (12 - Val(DTPicker1.Month))
                DTPicker1.Year = 2004
                If DTPicker1.Month >= 4 And DTPicker1.Year = 2004 Then
                    DTPicker1.Month = 3
                    DTPicker1.Day = 31
                Else
                'DTPicker1.Year = DTPicker1.Year + 1
                End If
            End If
        End If
        nodays Format(DTPicker1.Value, "MM"), Format(DTPicker1.Value, "YYYY")
        DTPicker1.Day = dacount
        If DTPicker1.Month >= 10 Then
            DTPicker1.Month = 3
            DTPicker1.Day = 31
            .TextMatrix(A, 3) = Format(DTPicker1.Value, "DD/MMM/YYYY")
            .TextMatrix(A, 4) = Format(DTPicker1.Value, "DDDD")
        Else
        .TextMatrix(A, 3) = Format(DTPicker1.Value, "DD/MMM/YYYY")
        .TextMatrix(A, 4) = Format(DTPicker1.Value, "DDDD")
        End If
    Next
    For A = 1 To RN
        .TextMatrix(A, 5) = Val(Format(.TextMatrix(A, 1), "MM"))
    Next
    End If
    End With
End Sub

Private Sub scw()

    Dim r As Integer
    Dim C As Integer
    For r = 1 To MS.Rows - 1
        For C = 0 To MS.Cols - 1
        MS.TextMatrix(r, C) = ""
        Next
    Next
    
    MS.ColWidth(0) = 600
    MS.ColWidth(1) = 1700
    MS.ColWidth(2) = 1325
    MS.ColWidth(3) = 1700
    MS.ColWidth(4) = 1325
    MS.ColWidth(5) = 1130
End Sub

