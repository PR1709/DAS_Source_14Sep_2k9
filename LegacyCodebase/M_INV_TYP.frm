VERSION 5.00
Begin VB.Form M_INV_TYP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Default Settings"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "M_INV_TYP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check27 
      Caption         =   "&5 Report Converted Optional RRW Format"
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Tag             =   "Write Reason Here!"
      Top             =   4680
      Width           =   3375
   End
   Begin VB.TextBox M_EML_IDY 
      Appearance      =   0  'Flat
      Height          =   425
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   5055
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   5520
      TabIndex        =   4
      Top             =   5040
      Width           =   2775
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   220
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   220
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Check to Enable Default Settings"
      Height          =   4000
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   8055
      Begin VB.CheckBox Check26 
         Caption         =   "&4 Global Customer Code - GCC Enable"
         Height          =   255
         Left            =   4080
         TabIndex        =   36
         Tag             =   "Write Reason Here!"
         Top             =   3480
         Width           =   3135
      End
      Begin VB.CheckBox Check25 
         Caption         =   "&3 Network Login - Improper ShutDown"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Tag             =   "Write Reason Here!"
         Top             =   3480
         Width           =   3135
      End
      Begin VB.CheckBox Check24 
         Caption         =   "&2 No Open Market Schemes to HVOs"
         Height          =   255
         Left            =   4080
         TabIndex        =   34
         Tag             =   "Write Reason Here!"
         Top             =   3240
         Width           =   3135
      End
      Begin VB.CheckBox Check23 
         Caption         =   "&1 Apply SKU Level Disc AfterTax - Sales"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Tag             =   "Write Reason Here!"
         Top             =   3240
         Width           =   3735
      End
      Begin VB.CheckBox Check6 
         Caption         =   "&E Invoice Entry in Cases"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox Check5 
         Caption         =   "&F System ID for Invoice Entry"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox Check4 
         Caption         =   "&D Van Load Out Reference while Invoice Entry"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "&C Direct Print after Transaction Save "
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         Caption         =   "&B Disable Stock Update while Invoice Entry"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   620
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&A Batch Mfg and Expiry Date"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox Check7 
         Caption         =   "&G Refresh Indexes before Exit"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   25
         Text            =   "7"
         Top             =   3000
         Width           =   615
      End
      Begin VB.CheckBox Check22 
         Caption         =   "&Z Txn Entry Days Since Last DayClosure"
         Height          =   255
         Left            =   4080
         TabIndex        =   24
         Tag             =   "Write Reason Here!"
         Top             =   3000
         Width           =   3255
      End
      Begin VB.CheckBox Check21 
         Caption         =   "&Y Rate B  for Van Challan Ref Invoice Entry"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Tag             =   "Write Reason Here!"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.CheckBox Check20 
         Caption         =   "&W Beat Plan Route for Invoice Entry"
         Height          =   255
         Left            =   4080
         TabIndex        =   22
         Tag             =   "Write Reason Here!"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.CheckBox Check19 
         Caption         =   "&T Show Activity Log at Startup"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Tag             =   "Write Reason Here!"
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   20
         Text            =   "7"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:mm:ss AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   3125
         TabIndex        =   19
         Text            =   "0:30"
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Check18 
         Caption         =   "&R Set Future Transaction Days"
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   2520
         Width           =   2535
      End
      Begin VB.CheckBox Check17 
         Caption         =   "&Q Set Session Idle Time"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CheckBox Check16 
         Caption         =   "&P Disable Report Date Control Settings"
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CheckBox Check15 
         Caption         =   "&O Mask Day Closure Records - ID Select View"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   3615
      End
      Begin VB.CheckBox Check14 
         Caption         =   "&N AutoBackup Before Exit - Keep Checked !"
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         Top             =   1800
         Width           =   3735
      End
      Begin VB.CheckBox Check13 
         Caption         =   "&M Glass Excess by Half Crate Increment Shell"
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   1560
         Width           =   3735
      End
      Begin VB.CheckBox Check12 
         Caption         =   "&L Auto Apply Schemes while Invoice Entry"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   1320
         Width           =   3735
      End
      Begin VB.CheckBox Check11 
         Caption         =   "&K Tax Incentive Location - VAT Calculation"
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CheckBox Check10 
         Caption         =   "&J Edit Van Reference while Invoice Modification "
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   840
         Width           =   3735
      End
      Begin VB.CheckBox Check9 
         Caption         =   "&I Secure Invoice Modification till Previous Day"
         Height          =   255
         Left            =   4080
         TabIndex        =   2
         Top             =   620
         Width           =   3735
      End
      Begin VB.CheckBox Check8 
         Caption         =   "&H Enable Auto Round off"
         Height          =   255
         Left            =   4080
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8040
         Y1              =   2160
         Y2              =   2160
      End
   End
   Begin VB.ComboBox txtdirmsg 
      Height          =   315
      ItemData        =   "M_INV_TYP.frx":0442
      Left            =   240
      List            =   "M_INV_TYP.frx":0491
      TabIndex        =   0
      Text            =   "txtdirmsg"
      Top             =   480
      Width           =   8055
   End
   Begin VB.Label Label2 
      Caption         =   "Default eMail ID for Fast Backup or  Data Xfer"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Invoice Template Description for Direct Print After SAVE"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   3960
   End
End
Attribute VB_Name = "M_INV_TYP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_KeyPress(K As Integer)
If K = 13 Then Check2.SetFocus
End Sub
Private Sub Check11_KeyPress(K As Integer)
If K = 13 Then Check12.SetFocus
End Sub
Private Sub Check10_KeyPress(K As Integer)
If K = 13 Then Check11.SetFocus
End Sub
Private Sub Check12_KeyPress(K As Integer)
If K = 13 Then Check13.SetFocus
End Sub
Private Sub Check13_KeyPress(K As Integer)
If K = 13 Then Check14.SetFocus
End Sub
Private Sub Check14_KeyPress(K As Integer)
If K = 13 Then Check15.SetFocus
End Sub
Private Sub Check15_KeyPress(K As Integer)
If K = 13 Then Check16.SetFocus
End Sub
Private Sub Check16_KeyPress(K As Integer)
If K = 13 Then Check17.SetFocus
End Sub
Private Sub Check17_KeyPress(K As Integer)
If K = 13 Then text1.SetFocus
End Sub
Private Sub Check18_KeyPress(K As Integer)
If K = 13 Then Text2.SetFocus
End Sub
Private Sub Check19_KeyPress(K As Integer)
If K = 13 Then Check20.SetFocus
End Sub
Private Sub CHECK2_KeyPress(K As Integer)
If K = 13 Then Check3.SetFocus
End Sub
Private Sub Check20_KeyPress(K As Integer)
If K = 13 Then Check21.SetFocus
End Sub
Private Sub Check21_KeyPress(K As Integer)
If K = 13 Then Check22.SetFocus
End Sub
Private Sub Check22_KeyPress(K As Integer)
If K = 13 Then Text3.SetFocus
End Sub
Private Sub Check23_KeyPress(K As Integer)
If K = 13 Then Check24.SetFocus
End Sub
Private Sub Check24_KeyPress(K As Integer)
If K = 13 Then Check25.SetFocus
End Sub
Private Sub check25_keypress(K As Integer)
If K = 13 Then Check26.SetFocus
End Sub
Private Sub check26_keypress(K As Integer)
If K = 13 Then Check27.SetFocus
End Sub
Private Sub check27_keypress(K As Integer)
If K = 13 Then Command1.SetFocus
End Sub
Private Sub TEXT2_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
Check19.SetFocus
ElseIf K = 38 Then
Check18.SetFocus
End If
End Sub
Private Sub Text3_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub Text3_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Check23.SetFocus
    'Command1.SetFocus
ElseIf K = 38 Then Check23.SetFocus
End If
End Sub
Private Sub text3_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub
Private Sub CHECK3_KeyPress(K As Integer)
If K = 13 Then Check4.SetFocus
End Sub
Private Sub Check4_keypress(K As Integer)
If K = 13 Then Check6.SetFocus 'Modified by PNR
End Sub
Private Sub Check6_keypress(K As Integer)
If K = 13 Then Check5.SetFocus
End Sub
Private Sub Check5_keypress(K As Integer)
If K = 13 Then Check7.SetFocus
End Sub
Private Sub Check7_KeyPress(K As Integer)
If K = 13 Then Check8.SetFocus
End Sub
Private Sub Check8_KeyPress(K As Integer)
If K = 13 Then Check9.SetFocus
End Sub
Private Sub Check9_KeyPress(K As Integer)
If K = 13 Then Check10.SetFocus
End Sub
Private Sub Command1_Click()
save
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
Dim rsdir As New ADODB.Recordset
rsdir.Open "select * from dir where dir_idy like 'AUT' AND DIR_IDX LIKE 'IND'", con, adOpenKeyset
If rsdir.RecordCount > 0 Then txtdirmsg = Trim(rsdir("dir_msg"))
rsdir.Close
rsdir.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'EML' AND DIR_IDX LIKE 'IDY'", con, adOpenKeyset, adLockPessimistic
If rsdir.RecordCount > 0 Then M_EML_IDY = Trim(rsdir("DIR_MSG"))
rsdir.Close
rsdir.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'SSS'", con, adOpenKeyset, adLockPessimistic
If rsdir.RecordCount > 0 Then text1 = Trim(rsdir("DIR_ALS"))
rsdir.Close
rsdir.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'FTD'", con, adOpenKeyset, adLockPessimistic
If rsdir.RecordCount > 0 Then Text2 = Trim(rsdir("DIR_ALS"))
rsdir.Close
rsdir.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DCD'", con, adOpenKeyset, adLockPessimistic
If rsdir.RecordCount > 0 Then Text3 = Trim(rsdir("DIR_ALS"))
rsdir.Close

If querycheck("DIR", "DIR_MSG", "E", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'EXP'") = True Then Check1.Value = 1
If querycheck("DIR", "DIR_MSG", "U", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'STK'") = True Then Check2.Value = 1
If querycheck("DIR", "DIR_MSG", "D", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'") = True Then Check3.Value = 1
If querycheck("DIR", "DIR_MSG", "V", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'VAN'") = True Then Check4.Value = 1
If querycheck("DIR", "DIR_MSG", "C", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'PCS'") = True Then Check6.Value = 1
If querycheck("DIR", "DIR_MSG", "N", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'NUM'") = True Then Check5.Value = 1
If querycheck("DIR", "DIR_ALS", "I", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'") = True Then Check7.Value = 1
If querycheck("DIR", "DIR_MSG", "R", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'ROU'") = True Then Check8.Value = 1
If querycheck("DIR", "DIR_MSG", "M", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'MOD'") = True Then Check9.Value = 1
If querycheck("DIR", "DIR_MSG", "F", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'REF'") = True Then Check10.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'INC'") = True Then Check11.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AAS'") = True Then Check12.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AIS'") = True Then Check13.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'ABK'") = True Then Check14.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'MDC'") = True Then Check15.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AYR'") = True Then Check16.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'SSS'") = True Then Check17.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'FTD'") = True Then Check18.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'SAY'") = True Then Check19.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DBR'") = True Then Check20.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DBV'") = True Then Check21.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DCD'") = True Then Check22.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DPV'") = True Then Check23.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AOM'") = True Then Check24.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'NWL'") = True Then Check25.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'GCC'") = True Then Check26.Value = 1
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'RRW'") = True Then Check27.Value = 1
End Sub
Private Sub Text1_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub Text1_KeyPress(K As Integer)
If K = 13 Or K = 40 Then
Check18.SetFocus
ElseIf K = 38 Then Check17.SetFocus
End If
End Sub
Private Sub Text2_GotFocus()
SendKeys "{HOME}+{END}"
End Sub
Private Sub Text2_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
Else
K = 0
End If
End Sub

Private Sub TXTDIRMSG_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Valid Report Name"
End Sub
Private Sub txtdirmsg_KeYPRESS(K As Integer)
If K = 13 Then
    Check1.SetFocus
End If
End Sub
Public Sub save()
Dim rsdir As New ADODB.Recordset
con.Execute "UPDATE DIR SET DIR_MSG='" & Trim(txtdirmsg) & "' WHERE DIR_IDY LIKE 'AUT' AND DIR_IDX LIKE 'IND'"
If rsdir.State = 1 Then rsdir.Close
rsdir.Open "select * from rri", PNL, adOpenKeyset, adLockPessimistic
rsdir.Move 6
If Trim(txtdirmsg) <> "" Then
rsdir!RI_REPORT = Trim(txtdirmsg)
rsdir.Update
End If
rsdir.Close
If Check1.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG = '' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'EXP'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "EXP"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG = 'E' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'EXP'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "EXP"
        rsdir!DIR_MSG = "E"
        rsdir.Update
    End If
End If
If Check2.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'STK'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "STK"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='U' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'STK'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "STK"
        rsdir!DIR_MSG = "U"
        rsdir.Update
    End If
End If

If Check3.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "DOS"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='D' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DOS'", CTR
        If CTR = 0 Then
        If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "DOS"
            rsdir!DIR_MSG = "D"
            rsdir.Update
        End If
End If

If Check4.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'VAN'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "VAN"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='V' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'VAN'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "VAN"
            rsdir!DIR_MSG = "V"
            rsdir.Update
        End If
End If
If Check6.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'PCS'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "PCS"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='C' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'PCS'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "PCS"
            rsdir!DIR_MSG = "C"
            rsdir.Update
        End If
End If
If Check5.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'NUM'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "NUM"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='N' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'NUM'", CTR
    If CTR = 0 Then
        If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "NUM"
        rsdir!DIR_MSG = "N"
        rsdir.Update
    End If
End If

If Check7.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_ALS='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "IND"
        rsdir!DIR_MSG = "PROPER SHUTDOWN"
        rsdir!DIR_ALS = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_ALS='I' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'", CTR
    If CTR = 0 Then
        If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "IND"
        rsdir!DIR_MSG = "PROPER SHUTDOWN"
        rsdir!DIR_ALS = "I"
        rsdir.Update
    End If
End If

If Check8.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'ROU'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "ROU"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='R' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'ROU'", CTR
    If CTR = 0 Then
        If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "ROU"
        rsdir!DIR_MSG = "R"
        rsdir.Update
    End If
End If

If Check9.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'MOD'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "MOD"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='M' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'MOD'", CTR
    If CTR = 0 Then
        If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "MOD"
        rsdir!DIR_MSG = "M"
        rsdir.Update
    End If
End If

If Check10.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'REF'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "REF"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='F' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'REF'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "REF"
            rsdir!DIR_MSG = "F"
            rsdir.Update
        End If
End If
If Check11.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'INC'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "INC"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'INC'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "INC"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
If Check12.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AAS'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "AAS"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AAS'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "AAS"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
If Check13.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AIS'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "AIS"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AIS'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "AIS"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
If Check14.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'ABK'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "ABK"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'ABK'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "ABK"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
' Mask Day Closure Check Flag
If Check15.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'MDC'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "MDC"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'MDC'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "MDC"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
' Disable Report Date Controls for Any Year Reporting
If Check16.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AYR'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "AYR"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AYR'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "AYR"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
' Session Idle Time
If Check17.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'SSS'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "SSS"
        rsdir!DIR_ALS = LTrim(str(Val(text1)))
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'SSS'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "SSS"
            rsdir!DIR_ALS = LTrim(str(Val(text1)))
            rsdir!DIR_MSG = "T"
            rsdir.Update
        Else
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'SSS'", con, adOpenKeyset, adLockPessimistic
            If rsdir.RecordCount > 0 Then
                rsdir!DIR_ALS = LTrim(str(Val(text1)))
                rsdir.Update
            End If
        End If
End If
' Set Future Transaction Days
If Check18.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'FTD'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "FTD"
        rsdir!DIR_ALS = LTrim(str(Text2))
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'FTD'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "FTD"
            rsdir!DIR_ALS = LTrim(str(Text2))
            rsdir!DIR_MSG = "T"
            rsdir.Update
        Else
        If Val(Text2) < 0 Or Val(Text2) > 30 Then
            MsgBox "Invalid Future Txn Days Value, Valid Between 0 to 30 Days!", vbInformation, "DAS Version 4.0"
            Text2.SetFocus
        Else
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'FTD'", con, adOpenKeyset, adLockPessimistic
            If rsdir.RecordCount > 0 Then
                rsdir!DIR_ALS = LTrim(str(Val(Text2)))
                rsdir.Update
            End If
            End If
        End If
End If
' Show Activity Log at Startup
If Check19.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'SAY'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "SAY"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'SAY'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "SAY"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
' Default Invoice Route ie, to pickup from Van Route or from Beat Plan, if checked will pickup from Beat Plan otherwise from Van
If Check20.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DBR'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "DBR"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DBR'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "DBR"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
' Default Invoice B Rate Applicable for Van Challan Ref Invoices - Check Flag
If Check21.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DBV'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "DBV"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DBV'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "DBV"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
' *
' Default Value for Transactions Entry from Day Closure Date
If Check22.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DCD'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "DCD"
        rsdir!DIR_ALS = LTrim(str(Text3))
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DCD'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "DCD"
            rsdir!DIR_ALS = LTrim(str(Text3))
            rsdir!DIR_MSG = "T"
            rsdir.Update
        Else
        If Val(Text3) < 0 Or Val(Text3) > 90 Then
            MsgBox "Invalid Txn Entry Days' Day Closure Value, Valid Between 0 to 90 Days!", vbInformation, "DAS Version 4.0"
            Text3.SetFocus
        Else
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DCD'", con, adOpenKeyset, adLockPessimistic
            If rsdir.RecordCount > 0 Then
                rsdir!DIR_ALS = LTrim(str(Val(Text3)))
                rsdir.Update
            End If
            End If
        End If
End If
' *
' Toggle for Tax Calculation at the time of Invoice / Sales Order
If Check23.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DPV'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "DPV"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'DPV'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "DPV"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
' *
' Default Value for Transactions Entry from Day Closure Date
If Check24.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AOM'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "AOM"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AOM'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "AOM"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
' *
' Setting for Network Login after Improper Shutdown in Network Systems
If Check25.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'NWL'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "NWL"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'NWL'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "NWL"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
' *
' Setting for Global Customer Code GCC for Retail Outlets
If Check26.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'GCC'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "GCC"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'GCC'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "GCC"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
' *
' Setting for Reports in RRW Format
If Check27.Value = 0 Then
    con.Execute "UPDATE DIR SET DIR_MSG='' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'RRW'", CTR
    If CTR = 0 Then
    If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "SYS"
        rsdir!DIR_IDX = "RRW"
        rsdir!DIR_MSG = ""
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG='T' WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'RRW'", CTR
        If CTR = 0 Then
            If rsdir.State = 1 Then rsdir.Close
            rsdir.Open "select * from dir ", con, adOpenKeyset, adLockPessimistic
            rsdir.AddNew
            rsdir!DIR_IDY = "SYS"
            rsdir!DIR_IDX = "RRW"
            rsdir!DIR_MSG = "T"
            rsdir.Update
        End If
End If
' *
If Trim(M_EML_IDY) <> "" Then
    If rsdir.State = 1 Then rsdir.Close
    rsdir.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'EML' AND DIR_IDX LIKE 'IDY'", con, adOpenKeyset, adLockPessimistic
    If rsdir.EOF = True Then
        rsdir.AddNew
        rsdir!DIR_IDY = "EML"
        rsdir!DIR_IDX = "IDY"
        rsdir!DIR_MSG = Trim(LCase(M_EML_IDY))
        rsdir!DIR_ALS = ""
        rsdir!DIR_DAT = Date
        rsdir.Update
    Else
        rsdir.MoveFirst
        rsdir!DIR_MSG = Trim(LCase(M_EML_IDY))
        rsdir.Update
    End If
Else
    con.Execute "UPDATE DIR SET DIR_MSG = 'effmcg@consoftsystems.com' WHERE DIR_IDY LIKE 'EML' AND DIR_IDX LIKE 'IDY'", CTR
    If CTR = 0 Then
        If rsdir.State = 1 Then rsdir.Close
        rsdir.Open "SELECT * FROM DIR", con, adOpenKeyset, adLockPessimistic
        rsdir.AddNew
        rsdir!DIR_IDY = "EML"
        rsdir!DIR_IDX = "IDY"
        rsdir!DIR_MSG = "effmcg@consoftsystems.com"
        rsdir!DIR_ALS = ""
        rsdir!DIR_DAT = Date
        rsdir.Update
    End If
End If
'Unload Me
End Sub
