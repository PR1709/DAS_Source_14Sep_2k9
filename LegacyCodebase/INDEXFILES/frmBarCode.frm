VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBarCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SKU Bar Code Print Interface"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   ControlBox      =   0   'False
   Icon            =   "frmBarCode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Disp All"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtNos 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1758
      MaxLength       =   4
      TabIndex        =   5
      Top             =   4680
      Width           =   975
   End
   Begin MSComctlLib.ListView lvBatch 
      Height          =   2415
      Left            =   2880
      TabIndex        =   4
      Top             =   1920
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Batch Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pur Rate"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sell Rate"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "MRP"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   315
      Left            =   2610
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   3495
   End
   Begin VB.TextBox txtBarCodeTemplate 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   240
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog dlgPrinter 
      Left            =   3840
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Printer"
   End
   Begin VB.CommandButton cmdBarCodeTemplate 
      Caption         =   "..."
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin MSComDlg.CommonDialog dlgTemplate 
      Left            =   3360
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Bar Code Template"
      Filter          =   "*.btw"
      InitDir         =   "C:\Program Files\EFFMCG4"
   End
   Begin VB.ListBox lstcur_lev 
      Height          =   2790
      Left            =   9000
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox lstmrp_rat 
      Height          =   2790
      Left            =   8880
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox lstprd_idy 
      Height          =   2790
      Left            =   9240
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox lstprd_nme 
      Height          =   2010
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label lblProcess 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nos"
      Height          =   195
      Left            =   1200
      TabIndex        =   17
      Top             =   4770
      Width           =   285
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      BorderStyle     =   6  'Inside Solid
      X1              =   0
      X2              =   7800
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblPartName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4560
      TabIndex        =   16
      Top             =   1320
      Width           =   2565
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Batches List for "
      Height          =   195
      Left            =   2880
      TabIndex        =   15
      Top             =   1320
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Code Template (.BTW) File"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   330
      Width           =   2205
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Code Printer Name"
      Height          =   195
      Left            =   825
      TabIndex        =   13
      Top             =   840
      Width           =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   0
      X2              =   7800
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SKU - Product Name"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1485
   End
End
Attribute VB_Name = "frmBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FORMSTATUS As Integer
Dim rec As ADODB.Recordset

Private Sub Check1_Click()
    Set rec = New ADODB.Recordset
    Dim SqlStatement As String
    SqlStatement = "select prd_nme,prd_idy,mrp_rat,cur_lev from prd WHERE STK_TYP = 'Salable' "
    If Check1.Value = 1 Then
        SqlStatement = SqlStatement & " and CUR_LEV > 0 "
    End If
    SqlStatement = SqlStatement & "  order by prd_idy "
    rec.Open SqlStatement, CON, adOpenForwardOnly, adLockReadOnly

    lstprd_nme.CLEAR
    lstmrp_rat.CLEAR
    lstprd_idy.CLEAR
    lstcur_lev.CLEAR
    While Not rec.EOF
        lstprd_nme.AddItem rec("prd_nme")
        lstprd_idy.AddItem rec("prd_idy")
        lstmrp_rat.AddItem rec("mrp_rat")
        lstcur_lev.AddItem rec("cur_lev")
        rec.MoveNext
    Wend
    If lstprd_nme.ListCount > 0 Then
        lstprd_nme.ListIndex = -1
    End If
    rec.Close

End Sub

'Dim con As Connection
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyE And Shift = vbAltMask Then
    cmdBarCodeTemplate_Click
End If
If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Activate()
DoEvents
    Dim status As Boolean
    Dim X As printer
    status = False
    Set rec = New ADODB.Recordset
    rec.Open "select prd_nme,prd_idy,mrp_rat,cur_lev from prd WHERE STK_TYP = 'Salable' and CUR_LEV > 0 order by prd_idy", CON, adOpenForwardOnly, adLockReadOnly
    lstprd_nme.CLEAR
    lstmrp_rat.CLEAR
    lstprd_idy.CLEAR
    lstcur_lev.CLEAR
    While Not rec.EOF
        lstprd_nme.AddItem rec("prd_nme")
        lstprd_idy.AddItem rec("prd_idy")
        lstmrp_rat.AddItem rec("mrp_rat")
        lstcur_lev.AddItem rec("cur_lev")
        rec.MoveNext
    Wend
    If lstprd_nme.ListCount > 0 Then
        lstprd_nme.ListIndex = -1
    End If
    rec.Close
    cmbPrinter.CLEAR
    
    For Each X In Printers
     cmbPrinter.AddItem X.DeviceName
    Next
    If cmbPrinter.ListCount < 1 Then
       MsgBox "No default printers installed"
    Else
        For I = 0 To cmbPrinter.ListCount - 1
            If cmbPrinter.List(I) = "TSC TTP-244 Plus" Then
                cmbPrinter.ListIndex = I
                status = True
                Exit For
            End If
        Next
    End If
    If status = False Then
        MsgBox "Bar Code Printer Not Installed", vbCritical, "DAS Version 4.0"
        Unload Me
    End If
    lblProcess.Visible = False
End Sub

Private Sub cmdBarCodeTemplate_Click()
  dlgTemplate.CancelError = True
  On Error GoTo ErrHandler
  dlgTemplate.Flags = cdlOFNHideReadOnly
  dlgTemplate.filter = "Bar Code Template (*.btw)|*.btw"
  dlgTemplate.FilterIndex = 2
  dlgTemplate.ShowOpen
  If dlgTemplate.fileName <> "" Then
    txtBarCodeTemplate.Text = dlgTemplate.fileName
  End If
ErrHandler:
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub lstprd_nme_Click()
    If lstprd_nme.ListIndex < 0 Then Exit Sub
    Dim Item As ListItem
    lvBatch.ListItems.CLEAR
    lstprd_idy.ListIndex = lstprd_nme.ListIndex
    lstmrp_rat.ListIndex = lstprd_nme.ListIndex
    lstcur_lev.ListIndex = lstprd_nme.ListIndex
    lblPartName.Caption = lstprd_nme.Text
    rec.Open "select bat_idy, prd_pdr, prd_sdr, prd_mrp, prd_qty from brd where prd_idy = '" & Trim(lstprd_idy.Text) & "'", CON, adOpenKeyset, adLockPessimistic
    While Not rec.EOF
        Set Item = lvBatch.ListItems.Add(, , rec("bat_idy"))
        Item.ListSubItems.Add , , DecimalValue(rec("prd_pdr"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("prd_sdr"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("prd_mrp"), 2)
        
        
       If IsNull(rec("PRD_QTY")) = True Then
            Item.ListSubItems.Add , , DecimalValue("0", 3)
       Else
            Item.ListSubItems.Add , , DecimalValue(rec("prd_qty"), 3)
       End If
        rec.MoveNext
    Wend
    rec.Close

End Sub

Private Sub lvBatch_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtNos.Text = Item.ListSubItems(4)
    If CInt(txtNos.Text) < 1 Then
      txtNos.Text = ""
    End If
End Sub

Private Sub txtNos_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And (KeyAscii < 47 Or KeyAscii > 58) Then KeyAscii = 0
End Sub

Private Sub txtSearch_Change()
    Dim status As Boolean
    status = False
    For I = 0 To lstprd_nme.ListCount - 1
        If Mid(RTrim(lstprd_nme.List(I)), 1, Len(txtSearch.Text)) = UCase(txtSearch.Text) Then
            lstprd_nme.ListIndex = I
            status = True
            Exit For
        End If
    Next
    If Not status Then
        lstprd_nme.ListIndex = -1
        lstprd_idy.ListIndex = -1
        lstmrp_rat.ListIndex = -1
        lstcur_lev.ListIndex = -1
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        lstprd_nme.SetFocus
    End If
End Sub

Private Sub Command1_Click()
    If txtBarCodeTemplate.Text = "" Then
        MsgBox "Please select the Bar Code Template file"
        cmdBarCodeTemplate.SetFocus
        Exit Sub
    End If
    If cmbPrinter.ListCount < 1 Then
       MsgBox "No default printers installed"
       cmbPrinter.SetFocus
       Exit Sub
    End If
    If Trim(txtNos.Text) = "" Then
        MsgBox "Please enter the no of labels"
        txtNos.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtNos.Text) Then
        MsgBox "Please enter the no of labels"
        txtNos.Text = ""
        txtNos.SetFocus
        Exit Sub
    End If
    If CInt(txtNos.Text) < 1 Then
        MsgBox "Please enter the no of labels"
        txtNos.Text = ""
        txtNos.SetFocus
        Exit Sub
    End If
    If MsgBox("Print " & txtNos.Text & " labels for    " & lstprd_nme.Text, vbYesNo + vbQuestion, "Bar Code") = vbNo Then
        Exit Sub
    End If
    Dim btPrintSetup As BARTENDER.PrintSetup
    Dim btApp As BARTENDER.Application
    Dim BTFormat As New BARTENDER.Format
    Dim btPrinterCodeTemplate As BARTENDER.PrinterCodeTemplate
    Dim btMessages As BARTENDER.Messages
    Dim ExportSuccess As Boolean
    Dim btMessageText As String
    Dim btPrintMethod As BARTENDER.PrintMethod
    Dim I As Integer
    Set btApp = New BARTENDER.Application   ' CreateObject("BarTender.Application")
    btApp.Visible = True
    Set BTFormat = btApp.Formats.Open(txtBarCodeTemplate.Text)
    lstprd_idy.ListIndex = lstprd_nme.ListIndex
    lstmrp_rat.ListIndex = lstprd_nme.ListIndex
    BTFormat.SetNamedSubStringValue "PRDIDY", lstprd_idy.Text
    BTFormat.SetNamedSubStringValue "MRP_RAT", lvBatch.SelectedItem.SubItems(3)
    BTFormat.SetNamedSubStringValue "PRD_PRN", lstprd_nme.Text
    Set btPrintSetup = BTFormat.PrintSetup
    btPrintSetup.printer = cmbPrinter.Text
    btPrintSetup.IdenticalCopiesOfLabel = CInt(txtNos.Text)
    BTFormat.save
    BTFormat.PrintOut
    BTFormat.Close (BARTENDER.BtSaveOptions.btSaveChanges)
    btApp.Quit (BARTENDER.BtSaveOptions.btDoNotSaveChanges)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Function DecimalValue(Value As String, DecimalPlaces As Integer) As String
    If IsNull(Value) Then Value = "0"
    'If Value Is Nothing Then Value = "0"
    If Value = "" Then Value = "0"
    If IsNumeric(Value) = False Then
        Value = "0"
    End If
    Value = CStr(Format(Value, IIf(DecimalPlaces = 3, "##########0.000", "##########0.00")))
    DecimalValue = Value
End Function

