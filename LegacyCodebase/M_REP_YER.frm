VERSION 5.00
Begin VB.Form M_REP_YER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Report Filters"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "M_REP_YER.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   3015
      Begin VB.OptionButton JAN 
         Caption         =   "Q&4 (Jan +Feb+Mar)"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   2400
         Width           =   1695
      End
      Begin VB.OptionButton APR 
         Caption         =   "Q&1 (Apr+May+June)"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton JUL 
         Caption         =   "Q&2 (July + Aug+ Sept)"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton OCT 
         Caption         =   "Q&3 (Oct+Nov+Dec)"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "M_REP_YER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim REC_SET_P As New ADODB.Recordset
Dim REC_SET_L As New ADODB.Recordset
Dim RS As New ADODB.Recordset
If Me.Tag = "UNT" Then
Me.MousePointer = vbHourglass
    REC_SET_P.Open "SELECT PRD_IDY, CUR_LEV, SUB_UNT FROM PRD", con, adOpenKeyset, adLockOptimistic
    If REC_SET_P.EOF = False Then REC_SET_P.MoveFirst
    REC_SET_L.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockOptimistic
    While REC_SET_P.EOF = False
           If REC_SET_L.EOF = True Then
           Else
           REC_SET_L.MoveFirst
           REC_SET_L.Find "PRD_IDY = '" & REC_SET_P("PRD_IDY") & "'"
           End If
           If REC_SET_L.EOF = True Then
              REC_SET_L.AddNew
              REC_SET_L("PRD_IDY") = REC_SET_P("PRD_IDY")
              REC_SET_L("PUR_STK") = 0
              REC_SET_L("OPN_STK") = 0
              REC_SET_L("PRT_STK") = 0
              REC_SET_L(4) = 0
              REC_SET_L(5) = 0
              REC_SET_L(6) = 0
              REC_SET_L(7) = 0
              REC_SET_L(8) = 0
            Else
              REC_SET_L("PUR_STK") = 0
              REC_SET_L("OPN_STK") = 0
              REC_SET_L("PRT_STK") = 0
            End If
            REC_SET_L.Update
        REC_SET_P.MoveNext
    Wend
    
    If APR.Value = True Then
        REC_SET_L.Requery
        For A = 4 To 6
            If RS.State = 1 Then RS.Close
            RS.Open "SELECT SUM(((PRD_QTY-INT(PRD_QTY))*1000)+(INT(PRD_QTY)*SUB_UNT)) TOTAL,INV.PRD_IDY FROM INV,PRD WHERE PRD.PRD_IDY LIKE INV.PRD_IDY  AND TXN_TYP IN('LSL','SAL','FRS','FRC') AND MONTH(INV_DAT)=" & A & " GROUP BY INV.PRD_IDY", con, adOpenKeyset, adLockOptimistic
            While RS.EOF = False
                REC_SET_L.Find "PRD_IDY LIKE '" & Trim(RS("PRD_IDY")) & "'"
                If REC_SET_L.EOF = False Then
                    If A = 4 Then
                        REC_SET_L.Update "OPN_STK", Val(RS(0) & "")
                    ElseIf A = 5 Then
                        REC_SET_L.Update "PUR_STK", Val(RS(0) & "")
                    ElseIf A = 6 Then
                        REC_SET_L.Update "PRT_STK", Val(RS(0) & "")
                    End If
                End If
                RS.MoveNext
                REC_SET_L.MoveFirst
            Wend
            REC_SET_L.Requery
        Next
        Me.MousePointer = vbDefault
        Openreport CURDIR & "\mir.rp1", "Products (Qty) Record Q2", 0, "", 6
    ElseIf JAN.Value = True Then
        For A = 1 To 3
            If RS.State = 1 Then RS.Close
            RS.Open "SELECT SUM(((PRD_QTY-INT(PRD_QTY))*1000)+(INT(PRD_QTY)*SUB_UNT)) TOTAL,INV.PRD_IDY FROM INV,PRD WHERE PRD.PRD_IDY LIKE INV.PRD_IDY  AND TXN_TYP IN('LSL','SAL','FRS','FRC') AND MONTH(INV_DAT)=" & A & " GROUP BY INV.PRD_IDY", con, adOpenKeyset, adLockOptimistic
            While RS.EOF = False
                REC_SET_L.Find "PRD_IDY LIKE '" & Trim(RS("PRD_IDY")) & "'"
                If REC_SET_L.EOF = False Then
                    If A = 1 Then
                        REC_SET_L.Update "OPN_STK", Val(RS(0) & "")
                    ElseIf A = 2 Then
                        REC_SET_L.Update "PUR_STK", Val(RS(0) & "")
                    ElseIf A = 3 Then
                        REC_SET_L.Update "PRT_STK", Val(RS(0) & "")
                    End If
                End If
                RS.MoveNext
                REC_SET_L.MoveFirst
            Wend
        Next
        Me.MousePointer = vbDefault
        Openreport CURDIR & "\mir.rp1", "Products (Qty) Record Q1", 0, "", 6
    ElseIf JUL.Value = True Then
        For A = 7 To 9
            If RS.State = 1 Then RS.Close
            RS.Open "SELECT SUM(((PRD_QTY-INT(PRD_QTY))*1000)+(INT(PRD_QTY)*SUB_UNT)) TOTAL,INV.PRD_IDY FROM INV,PRD WHERE PRD.PRD_IDY LIKE INV.PRD_IDY  AND TXN_TYP IN('LSL','SAL','FRS','FRC') AND MONTH(INV_DAT)=" & A & " GROUP BY INV.PRD_IDY", con, adOpenKeyset, adLockOptimistic
            While RS.EOF = False
                REC_SET_L.Find "PRD_IDY LIKE '" & Trim(RS("PRD_IDY")) & "'"
                If REC_SET_L.EOF = False Then
                    If A = 7 Then
                        REC_SET_L.Update "OPN_STK", Val(RS(0) & "")
                    ElseIf A = 8 Then
                        REC_SET_L.Update "PUR_STK", Val(RS(0) & "")
                    ElseIf A = 9 Then
                        REC_SET_L.Update "PRT_STK", Val(RS(0) & "")
                    End If
                End If
                RS.MoveNext
                REC_SET_L.MoveFirst
                REC_SET_L.Requery
            Wend
        Next
        Me.MousePointer = vbDefault
        Openreport CURDIR & "\mir.rp1", "Products (Qty) Record Q3", 0, "", 6
    ElseIf OCT.Value = True Then
        For A = 10 To 12
            If RS.State = 1 Then RS.Close
            RS.Open "SELECT SUM(((PRD_QTY-INT(PRD_QTY))*1000)+(INT(PRD_QTY)*SUB_UNT)) TOTAL,INV.PRD_IDY FROM INV,PRD WHERE PRD.PRD_IDY LIKE INV.PRD_IDY  AND TXN_TYP IN('LSL','SAL','FRS','FRC') AND MONTH(INV_DAT)=" & A & " GROUP BY INV.PRD_IDY", con, adOpenKeyset, adLockOptimistic
            While RS.EOF = False
            REC_SET_L.Requery
                REC_SET_L.Find "PRD_IDY LIKE '" & Trim(RS("PRD_IDY")) & "'"
                If REC_SET_L.EOF = False Then
                    If A = 10 Then
                        REC_SET_L.Update "OPN_STK", Val(RS(0) & "")
                    ElseIf A = 11 Then
                        REC_SET_L.Update "PUR_STK", Val(RS(0) & "")
                    ElseIf A = 12 Then
                        REC_SET_L.Update "PRT_STK", Val(RS(0) & "")
                    End If
                End If
                RS.MoveNext
                REC_SET_L.MoveFirst
            Wend
        Next
        Me.MousePointer = vbDefault
        Openreport CURDIR & "\mir.rp1", "Products (Qty) Record Q4", 0, "", 6
    End If
ElseIf Me.Tag = "VAL" Then
    If JAN.Value = True Then
        Openreport CURDIR & "\mir.rp1", "Products (Value) Record Q1", 2, "INV->TXN_TYP ='LSL'.OR.INV->TXN_TYP = 'SAL'", 6
    ElseIf APR.Value = True Then
        Openreport CURDIR & "\mir.rp1", "Products (Value) Record Q2", 0, "INV->TXN_TYP ='LSL'.OR.INV->TXN_TYP = 'SAL'", 6
    ElseIf JUL.Value = True Then
        Openreport CURDIR & "\mir.rp1", "Products (Value) Record Q3", 0, "", 6
    ElseIf OCT.Value = True Then
        Openreport CURDIR & "\mir.rp1", "Products (Value) Record Q4", 0, "", 6
    End If
ElseIf Me.Tag = "DLR" Then
    If JAN.Value = True Then
        Openreport CURDIR & "\mir.rp1", "Dealers Yearly Performance Q4", 0, "", 6
    ElseIf APR.Value = True Then
        Openreport CURDIR & "\mir.rp1", "Dealers Yearly Performance Q1", 0, "", 6
    ElseIf JUL.Value = True Then
        Openreport CURDIR & "\mir.rp1", "Dealers Yearly Performance Q2", 0, "", 6
    ElseIf OCT.Value = True Then
        Openreport CURDIR & "\mir.rp1", "Dealers Yearly Performance Q3", 0, "", 6
    End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Keypress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub OCT_kEYPRESS(K As Integer)
If K = 13 Then Command1.SetFocus
End Sub
Private Sub APR_kEYPRESS(K As Integer)
If K = 13 Then Command1.SetFocus
End Sub
Private Sub JUL_kEYPRESS(K As Integer)
If K = 13 Then Command1.SetFocus
End Sub
Private Sub JAN_kEYPRESS(K As Integer)
If K = 13 Then Command1.SetFocus
End Sub
