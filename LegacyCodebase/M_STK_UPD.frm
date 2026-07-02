VERSION 5.00
Begin VB.Form M_FRM_UPD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Effmcg"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "M_STK_UPD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Update Stock...?"
      Height          =   195
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "M_FRM_UPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
          Dim REC_SET_I As New ADODB.Recordset
    Dim REC_SET_P As New ADODB.Recordset
    con.Execute "UPDATE PRD SET CUR_LEV = 0"
    REC_SET_I.Open "SELECT * FROM INV", con, adOpenKeyset, adLockReadOnly
    REC_SET_P.Open "Select * FROM PRD", con, adOpenKeyset, adLockReadOnly
    REC_SET_I.Move REC_POS
    Do While True
        If REC_SET_I.EOF = True Then
            Timer1.Enabled = False
            Exit Sub
        Else
        'For A = 1 To REC_SET_I.RecordCount
            REC_SET_P.MoveFirst
            REC_SET_P.Find "PRD_IDY = '" & REC_SET_I.Fields("PRD_IDY") & "'"
            If REC_SET_P.EOF = False Then
                PRD_SUB = P_SUB_UNT(QTY(REC_SET_P("CUR_LEV") & ""), REC_SET_P("SUB_UNT"))
                INV_SUB = P_SUB_UNT(QTY(REC_SET_I("PRD_QTY") & ""), REC_SET_P("SUB_UNT"))
                If REC_SET_I("TXN_TYP") = "OPS" Or REC_SET_I("TXN_TYP") = "PUR" Or REC_SET_I("TXN_TYP") = "SRN" Or REC_SET_I("TXN_TYP") = "STO" Or REC_SET_I("TXN_TYP") = "RSR" Or REC_SET_I("TXN_TYP") = "PCR" Then
                    M_QTY = P_QTY_UNT(Val(PRD_SUB) + Val(INV_SUB), REC_SET_P("SUB_UNT"))
                    con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET_I("PRD_IDY")) & "' SET CUR_LEV = " & M_QTY
                    'REC_SET_P.Update
                Else
                    M_QTY = P_QTY_UNT(Val(PRD_SUB) - Val(INV_SUB), REC_SET_P("SUB_UNT"))
                    con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET_I("PRD_IDY")) & "' SET CUR_LEV = " & M_QTY
                    'MsgBox REC_SET_P(0)
                    'REC_SET_P.Update
                End If
            'End If
            REC_SET_I.MoveNext
            End If
        End If
        REC_POS = REC_POS + 1
        'If REC_POS Mod 1000 = 0 Then
            'MsgBox "WAIT"
            'Exit Sub
        'End If
        M_MDI.Caption = REC_POS
    Loop
    'Next
  
End Sub
