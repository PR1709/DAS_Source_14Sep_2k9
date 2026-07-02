VERSION 5.00
Begin VB.Form M_REP_IQY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Sales Option"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   Icon            =   "M_REP_IQY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox all 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "&All"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton check2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Se&c. Sales"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   1815
   End
   Begin VB.OptionButton check1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "&Sales"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "M_REP_IQY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INV_TYP As String

Private Sub Check4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check4.Value = 1 Then
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        Command1.SetFocus
    Else
        Check1.SetFocus
    End If
End If
End Sub

Private Sub all_Click()
If all.Value = 1 Then Command1_Click
End Sub

Private Sub Command1_Click()
If all.Value <> 1 Then
    If Check1.Value = True Then
    con.Execute "UPDATE DIR WHERE DIR_IDY LIKE 'DOC' AND DIR_IDX LIKE 'TYP' SET DIR_ALS= 'Cash'"
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='FRC')"
        Else
            SAL_QRY = "INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='FRC'"
        End If
    End If
    If Check2.Value = True Then
        con.Execute "UPDATE DIR WHERE DIR_IDY LIKE 'DOC' AND DIR_IDX LIKE 'TYP' SET DIR_ALS= 'Credit'"
        If Trim(Len(SAL_QRY)) <> 0 Then
            SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='FRS')"
        Else
            SAL_QRY = "INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='FRS'"
        End If
    End If
Else
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='FRS'.OR.INV->TXN_TYP='FRC')"
    con.Execute "UPDATE DIR WHERE DIR_IDY LIKE 'DOC' AND DIR_IDX LIKE 'TYP' SET DIR_ALS='All'"
End If
'    Dim rs As New ADODB.Recordset
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from dir WHERE DIR_IDY LIKE 'HDR' AND DIR_IDX LIKE 'DAT'", con, adOpenKeyset, adLockOptimistic
'    DAT1 = Format(rs("DIR_ALS"), "DD/MM/YYYY")
'    DAT2 = Format(rs("DIR_MSG"), "DD/MM/YYYY")
    'Openscopereport CURDIR & "\report.rp1", "Sales Report", 2, 2, Trim(DAT1), Trim(DAT2), Trim(SAL_QRY), 6
    'OpenDocument CURDIR & "\REPORT.RP1", "Sales Report", 2, Trim(rs("dir_als")), Trim(rs("dir_msg")), 6, CURDIR & "\inv.cdx,c,inv_DAT"
If M_REP_IQY.Tag = "Sales Register" Then
    Openreport CURDIR & "\report.rp1", "Sales Register", 2, Trim(SAL_QRY), 6
ElseIf M_REP_IQY.Tag = "Van Sales" Then
    Openreport CURDIR & "\report.rp1", "Sales Report(Van)", 2, Trim(SAL_QRY) & ".and.inl1->INV_idy='RSI0000003'", 6
End If
SAL_QRY = ""
Unload M_REP_IQY
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
checkdongle
Check1.SetFocus
Check1.Value = 1
Check2.Value = 0
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub Form_Keypress(k As Integer)
If k = 27 Then Unload Me
End Sub

Public Sub UPDATELSS()
    
    SAL_QRY = SAL_QRY & ".AND.(INV->TXN_TYP='LSL'.OR.INV->TXN_TYP='SAL'.OR.INV->TXN_TYP='FRS'.OR.INV->TXN_TYP='FRC')"
    con.Execute "UPDATE DIR WHERE DIR_IDY LIKE 'DOC' AND DIR_IDX LIKE 'TYP' SET DIR_ALS='All'"

    If all.Value = 1 Then
        INV_TYP = "'SAL','LSL','FRS','FRC'"
    Else
        If Check1.Value = True Then
            INV_TYP = "'LSL','FRC'"
        Else
            INV_TYP = "'SAL','FRS'"
        End If
    End If
    UPDATELSS

Me.MousePointer = vbHourglass
If LSS_TYP = "NUM" Then
SAL_QRY = SAL_QRY & ".AND.INL->ORD_IDY ='0'"
Else
    If "'SAL','LSL','FRS','FRC'" = INV_TYP Then
        SAL_QRY = "inv->inv_dat>=ctod('" & FRM_INV_DAT & "').and.inv->inv_dat<=ctod('" & TO_INV_DAT & "').AND.INL->ROU_IDY = '" & Trim(ROU_IDY) & "'.AND.(INV->TXN_TYP = 'SAL'.OR.INV->TXN_TYP = 'LSL'.OR.INV->TXN_TYP ='FRS'.OR.INV->TXN_TYP ='FRC')"
    ElseIf "'LSL','FRC'" = INV_TYP Then
        SAL_QRY = "inv->inv_dat>=ctod('" & FRM_INV_DAT & "').and.inv->inv_dat<=ctod('" & TO_INV_DAT & "').AND.INL->ROU_IDY = '" & Trim(ROU_IDY) & "'.AND.(INV->TXN_TYP = 'LSL'.OR.INV->TXN_TYP ='FRC')"
    ElseIf "'SAL','FRS'" = INV_TYP Then
        SAL_QRY = "inv->inv_dat>=ctod('" & FRM_INV_DAT & "').and.inv->inv_dat<=ctod('" & TO_INV_DAT & "').AND.INL->ROU_IDY = '" & Trim(ROU_IDY) & "'.AND.(INV->TXN_TYP = 'SAL'.OR.INV->TXN_TYP ='FRS')"
    End If
End If
Me.MousePointer = vbDefault
Openreport CURDIR & "\Report.rp1", "Loading Sheet", 2, SAL_QRY, 6
End Sub
