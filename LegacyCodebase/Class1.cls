VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub SaveCalender()
If CHECKTB("wek.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If
' Check for TWP Version NO SAVE
Dim DIR_RS As New ADODB.Recordset
If DIR_RS.State = 1 Then DIR_RS.Close
DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'MDL' AND DIR_IDX = 'TWP'"), CON, adOpenKeyset, adLockPessimistic
If DIR_RS.RecordCount > 0 Then
    If Trim(DIR_RS("DIR_ALS")) = "Y" Then
        MsgBox ("Cannot Save Master Data for KO-TWP Version!"), vbCritical, "DAS Version 4.0"
'        cancel
        Exit Sub
    End If
End If
'
If MsgBox("This Action Will Overwrite Existing Data! Continue?", vbOKCancel) = vbOK Then
CON.Execute "delete from tgt"
CON.Execute "delete from tgb"
CON.Execute "delete from tgc"
CON.Execute "delete from tgs"
Dim rs As New ADODB.Recordset
With M_FRM_CAL
'For A = 1 To .ms.Rows - 1
'    If Len(Trim(.ms.TextMatrix(A, 2))) = 0 Or Len(Trim(.ms.TextMatrix(A, 2))) = 0 Then
'        MsgBox "Cannot Save FMCG Calender", vbCritical
'        .SetFocus
'        Exit Sub
'    End If
'Next
'For A = 1 To .ms.Rows - 1
'    If DateDiff("D", .ms.TextMatrix(A, 1), .ms.TextMatrix(A, 3)) < 0 Then
'        MsgBox "Cannot Save FMCG Calender", vbCritical
'        .ms.Row = A
'        .ms.Col = 3
'        .ms.SetFocus
'        Exit Sub
'    End If
'Next
CON.Execute "DELETE FROM WEK"
rs.Open "SELECT * FROM WEK", CON, adOpenKeyset, adLockPessimistic
For A = 1 To .MS.Rows - 1
    If Len(Trim(.MS.TextMatrix(A, 1))) > 0 Or Len(Trim(.MS.TextMatrix(A, 2))) > 0 Then
        rs.AddNew
        rs(0) = Val(.MS.TextMatrix(A, 0) & "")
        rs(1) = .MS.TextMatrix(A, 1)
        rs(2) = .MS.TextMatrix(A, 3)
        rs(3) = Val(.MS.TextMatrix(A, 5))
        rs.Update
    End If
Next
End With
    Unload M_FRM_CAL
Else
    Unload M_FRM_CAL
End If
DoEvents
End Sub
Public Sub TGT_LOAD_SKU()
If CHECKTB("tgt.dbf,prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If
Dim rs As New ADODB.Recordset
With M_FRM_TGT
.MousePointer = vbHourglass
'If .SMNID = "All" Then
'    rs.Open "select * from TGT,PRD WHERE PRD.PRD_IDY LIKE TGT.PRD_IDY AND SMN_IDY LIKE '0'", CON, adOpenKeyset
'Else
    rs.Open "select * from TGT,PRD WHERE PRD.PRD_IDY LIKE TGT.PRD_IDY AND SMN_IDY LIKE '" & Trim(.smnid) & "'", CON, adOpenKeyset
'End If
.MS.Rows = 2
If rs.RecordCount > 0 Then
    .M_BRD.Enabled = True
    .M_SEG.Enabled = True
End If
For A = 1 To rs.RecordCount
    For d = 2 To .MS.Rows - 1
        If Trim(.MS.TextMatrix(d, 105)) = Trim(rs!prd_idy) Then
            .MS.TextMatrix(d, (Val(rs!WEK_IDY) * 2)) = Int(Val(rs!TGT_QTY))
            .MS.TextMatrix(d, (Val(rs!WEK_IDY) * 2) - 1) = Int(Val(rs!TGT_CFC))
            GoTo li
        End If
    Next
    .MS.AddItem ""
    .MS.TextMatrix(.MS.Rows - 1, 0) = Trim(rs!PRD_NME)
    .MS.TextMatrix(.MS.Rows - 1, 105) = Trim(rs!prd_idy)
    .MS.TextMatrix(.MS.Rows - 1, (Val(rs!WEK_IDY) * 2) - 1) = Int(Val(rs!TGT_CFC))
    .MS.TextMatrix(.MS.Rows - 1, Val(rs!WEK_IDY) * 2) = Int(Val(rs!TGT_QTY))
li:
    rs.MoveNext
Next

If rs.State = 1 Then rs.Close
'If .SMNID = "All" Then
'    rs.Open "SELECT * FROM PRD WHERE PRD_IDY NOT IN(select PRD_IDY from TGT WHERE SMN_IDY LIKE '0' GROUP BY PRD_IDY)", CON, adOpenKeyset, adlockpessimistic
'Else
    rs.Open "SELECT * FROM PRD WHERE PRD_IDY NOT IN(select PRD_IDY from TGT WHERE SMN_IDY LIKE '" & Trim(.smnid) & "' GROUP BY PRD_IDY) and stk_typ in('Salable','Free')", CON, adOpenKeyset, adLockPessimistic
'End If
For A = 1 To rs.RecordCount
    .MS.AddItem ""
    .MS.TextMatrix(.MS.Rows - 1, 0) = Trim(rs!PRD_NME)
    .MS.TextMatrix(.MS.Rows - 1, 105) = Trim(rs!prd_idy)
    rs.MoveNext
Next
.MousePointer = vbDefault
End With
End Sub

Public Sub TGT_SAVE()
With M_FRM_TGT
If MsgBox("Confirm Targets Save ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    .CANCELRECORD
    Exit Sub
End If
If CHECKTB("tgt.dbf,tgb.dbf,tgc.dbf,tgs.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

If MsgBox("This Action will overwrite existing TARGETS information?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If .SSTab1.TabEnabled(0) = True Then
    If Len(Trim(.MS.TextMatrix(2, 0))) = 0 Then
        MsgBox "Cannot Save", vbExclamation, "DAS Version 4.0"
        Exit Sub
    End If
    .MousePointer = vbHourglass
    CON.Execute "DELETE FROM TGC"
    CON.Execute "DELETE FROM TGB"
    CON.Execute "DELETE FROM TGS"
    'If .SMNID = "All" Then
        'CON.Execute "DELETE FROM TGT WHERE SMN_IDY LIKE '0'"
    'Else
        CON.Execute "DELETE FROM TGT WHERE SMN_IDY LIKE '" & Trim(.smnid) & "'"
    'End If
    rs.Open "SELECT * FROM TGT", CON, adOpenKeyset, adLockPessimistic
    For A = 1 To 104 'COLS
        For b = 2 To .MS.Rows - 1 'ROWS
            rs.AddNew
            rs!prd_idy = Trim(.MS.TextMatrix(b, 105))
            'If .SMNID = "All" Then
             '   rs!SMN_IDY = 0
            'Else
                rs!SMN_IDY = .smnid
            'End If
            rs!TGT_CFC = Val(.MS.TextMatrix(b, A))
            rs!TGT_QTY = Val(.MS.TextMatrix(b, A + 1))
            rs!WEK_IDY = Val(Mid(Trim(.MS.TextMatrix(0, A)), 2, 2))
            rs.Update
        Next
        A = A + 1
    Next
ElseIf .SSTab1.TabEnabled(1) = True Then
    If Len(Trim(.msb.TextMatrix(2, 0))) = 0 Then
        MsgBox "Cannot Save", vbExclamation
        Exit Sub
    End If
    .MousePointer = vbHourglass
    CON.Execute "DELETE FROM TGC"
    CON.Execute "DELETE FROM TGS"
    CON.Execute "DELETE FROM TGT"
    'If .smnid = "All" Then
    '    CON.Execute "DELETE FROM TGB WHERE SMN_IDY LIKE '0'"
    'Else
        CON.Execute "DELETE FROM TGB WHERE SMN_IDY LIKE '" & .smnid & "'"
    'End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM TGB", CON, adOpenKeyset, adLockPessimistic
    For A = 1 To 104 'COLS
        For b = 2 To .msb.Rows - 1 'ROWS
            rs.AddNew
            rs!BRD_NME = Trim(.msb.TextMatrix(b, 0))
            If .smnid = "All" Then
                rs!SMN_IDY = 0
            Else
                rs!SMN_IDY = .smnid
            End If
            rs!TGT_CFC = Val(.msb.TextMatrix(b, A))
            rs!TGT_QTY = Val(.msb.TextMatrix(b, A + 1))
            rs!WEK_IDY = Val(Mid(Trim(.msb.TextMatrix(0, A)), 2, 2))
            rs.Update
        Next
        A = A + 1
    Next
ElseIf .SSTab1.TabEnabled(2) = True Then
    If Len(Trim(.MSS.TextMatrix(2, 0))) = 0 Then
        MsgBox "Cannot Save", vbExclamation, "DAS Version 4.0"
        Exit Sub
    End If
    .MousePointer = vbHourglass
    CON.Execute "DELETE FROM TGC"
    CON.Execute "DELETE FROM TGB"
    CON.Execute "DELETE FROM TGT"
    'If .smnid = "All" Then
    '    CON.Execute "DELETE FROM TGS WHERE SMN_IDY LIKE '0'"
    'Else
        CON.Execute "DELETE FROM TGS WHERE SMN_IDY LIKE '" & .smnid & "'"
    'End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM TGS", CON, adOpenKeyset, adLockPessimistic
    For A = 1 To 104 'COLS
        For b = 2 To .MSS.Rows - 1 'ROWS
            rs.AddNew
            rs!pgr_seg = Trim(.MSS.TextMatrix(b, 0))
            If .smnid = "All" Then
                rs!SMN_IDY = 0
            Else
                rs!SMN_IDY = .smnid
            End If
            rs!TGT_CFC = Val(.MSS.TextMatrix(b, A))
            rs!TGT_QTY = Val(.MSS.TextMatrix(b, A + 1))
            rs!WEK_IDY = Val(Mid(Trim(.MSS.TextMatrix(0, A)), 2, 2))
            rs.Update
        Next
        A = A + 1
    Next
ElseIf .SSTab1.TabEnabled(3) = True Then
    If Len(Trim(.MSC.TextMatrix(2, 0))) = 0 Then
        MsgBox "Cannot Save", vbExclamation, "DAS Version 4.0"
        Exit Sub
    End If
    .MousePointer = vbHourglass
    CON.Execute "DELETE FROM TGT"
    CON.Execute "DELETE FROM TGB"
    CON.Execute "DELETE FROM TGS"
    'If .smnid = "All" Then
     '   CON.Execute "DELETE FROM TGC WHERE SMN_IDY LIKE '0'"
    'Else
        CON.Execute "DELETE FROM TGC WHERE SMN_IDY LIKE '" & .smnid & "'"
    'End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM TGC", CON, adOpenKeyset, adLockPessimistic
    For A = 1 To 104 'COLS
        For b = 2 To .MSC.Rows - 1 'ROWS
            rs.AddNew
            rs!pgt_cat = Trim(.MSC.TextMatrix(b, 0))
            If .smnid = "All" Then
                rs!SMN_IDY = 0
            Else
                rs!SMN_IDY = .smnid
            End If
            rs!TGT_CFC = Val(.MSC.TextMatrix(b, A))
            rs!TGT_QTY = Val(.MSC.TextMatrix(b, A + 1))
            rs!WEK_IDY = Val(Mid(Trim(.MSC.TextMatrix(0, A)), 2, 2))
            rs.Update
        Next
        A = A + 1
    Next
End If
.MousePointer = vbDefault
MsgBox "Targets Saved", vbExclamation
.CANCELRECORD
End With
End Sub

Public Sub TGT_LOD_BRD()
If CHECKTB("tgb.dbf,prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

Dim rs As New ADODB.Recordset
With M_FRM_TGT
.MousePointer = vbHourglass
'If .SMNID = "All" Then
 '   rs.Open "select * from TGB WHERE SMN_IDY LIKE '0'", CON, adOpenKeyset, adlockpessimistic
'Else
    rs.Open "select * from TGB WHERE SMN_IDY LIKE '" & Trim(.smnid) & "'", CON, adOpenKeyset
'End If
.msb.Rows = 2
For A = 1 To rs.RecordCount
    For d = 2 To .msb.Rows - 1
        If Trim(.msb.TextMatrix(d, 0)) = Trim(rs!BRD_NME) Then
            .msb.TextMatrix(d, (Val(rs!WEK_IDY) * 2)) = Int(Val(rs!TGT_QTY))
            .msb.TextMatrix(d, (Val(rs!WEK_IDY) * 2) - 1) = Int(Val(rs!TGT_CFC))
            GoTo li
        End If
    Next
    .msb.AddItem ""
    .msb.TextMatrix(.msb.Rows - 1, 0) = Trim(rs!BRD_NME)
    .msb.TextMatrix(.msb.Rows - 1, (Val(rs!WEK_IDY) * 2) - 1) = Int(Val(rs!TGT_CFC))
    .msb.TextMatrix(.msb.Rows - 1, Val(rs!WEK_IDY) * 2) = Int(Val(rs!TGT_QTY))
li:
rs.MoveNext
Next

If rs.State = 1 Then rs.Close
'If .SMNID = "All" Then
'    rs.Open "SELECT * FROM PRD WHERE BRD_NME NOT IN(select BRD_NME from TGB WHERE SMN_IDY LIKE '0' GROUP BY BRD_NME) GROUP BY BRD_NME", CON, adOpenKeyset, adlockpessimistic
'Else
    rs.Open "SELECT * FROM PRD WHERE BRD_NME NOT IN(select BRD_NME from TGB WHERE SMN_IDY LIKE '" & Trim(.smnid) & "' GROUP BY BRD_NME) GROUP BY BRD_NME", CON, adOpenKeyset, adLockPessimistic
'End If
For A = 1 To rs.RecordCount
    .msb.AddItem ""
    .msb.TextMatrix(.msb.Rows - 1, 0) = Trim(rs!BRD_NME)
    rs.MoveNext
Next
.MousePointer = vbDefault
End With
End Sub

Public Sub TGT_LOD_SEG()
If CHECKTB("tgs.dbf,pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

Dim rs As New ADODB.Recordset
With M_FRM_TGT
.MousePointer = vbHourglass
'If .SMNID = "All" Then
'    rs.Open "select * from TGS WHERE SMN_IDY LIKE '0'", CON, adOpenKeyset
'Else
    rs.Open "select * from TGS WHERE SMN_IDY LIKE '" & Trim(.smnid) & "'", CON, adOpenKeyset
'End If

.MSS.Rows = 2
If rs.RecordCount <> 0 Then
    .M_CAT.Enabled = True
End If
For A = 1 To rs.RecordCount
    For d = 2 To .MSS.Rows - 1
        If Trim(.MSS.TextMatrix(d, 0)) = Trim(rs!pgr_seg) Then
            .MSS.TextMatrix(d, (Val(rs!WEK_IDY) * 2)) = Int(Val(rs!TGT_QTY))
            .MSS.TextMatrix(d, (Val(rs!WEK_IDY) * 2) - 1) = Int(Val(rs!TGT_CFC))
            GoTo li
        End If
    Next
    .MSS.AddItem ""
    .MSS.TextMatrix(.MSS.Rows - 1, 0) = Trim(rs!pgr_seg)
    .MSS.TextMatrix(.MSS.Rows - 1, (Val(rs!WEK_IDY) * 2) - 1) = Int(Val(rs!TGT_CFC))
    .MSS.TextMatrix(.MSS.Rows - 1, Val(rs!WEK_IDY) * 2) = Int(Val(rs!TGT_QTY))
li:
rs.MoveNext
Next

If rs.State = 1 Then rs.Close
'If .SMNID = "All" Then
'    rs.Open "SELECT * FROM PGR WHERE GRP_CLS NOT IN(select PGR_SEG from TGS WHERE SMN_IDY LIKE '0') GROUP BY GRP_CLS", CON, adOpenKeyset, adlockpessimistic
'Else
    rs.Open "SELECT * FROM PGR WHERE GRP_CLS NOT IN(select PGR_SEG from TGS WHERE SMN_IDY LIKE '" & Trim(.smnid) & "') GROUP BY GRP_CLS", CON, adOpenKeyset, adLockPessimistic
'End If
For A = 1 To rs.RecordCount
    .MSS.AddItem ""
    .MSS.TextMatrix(.MSS.Rows - 1, 0) = Trim(rs!GRP_CLS)
    rs.MoveNext
Next
.MousePointer = vbDefault
End With
End Sub

Public Sub TGT_LOD_CAT()
If CHECKTB("tgc.dbf,pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

Dim rs As New ADODB.Recordset
With M_FRM_TGT
.MousePointer = vbHourglass
'If .C_SMN_IDY = "All" Then
'    rs.Open "select * from TGC WHERE SMN_IDY LIKE '0'", CON, adOpenKeyset
'Else
    rs.Open "select * from TGC WHERE SMN_IDY LIKE '" & Trim(.smnid) & "'", CON, adOpenKeyset
'End If

.MSC.Rows = 2

For A = 1 To rs.RecordCount
    For d = 2 To .MSC.Rows - 1
        If Trim(.MSC.TextMatrix(d, 0)) = Trim(rs!pgt_cat) Then
            .MSC.TextMatrix(d, (Val(rs!WEK_IDY) * 2)) = Int(Val(rs!TGT_QTY))
            .MSC.TextMatrix(d, (Val(rs!WEK_IDY) * 2) - 1) = Int(Val(rs!TGT_CFC))
            GoTo li
        End If
    Next
    .MSC.AddItem ""
    .MSC.TextMatrix(.MSC.Rows - 1, 0) = Trim(rs!pgt_cat)
    .MSC.TextMatrix(.MSC.Rows - 1, (Val(rs!WEK_IDY) * 2) - 1) = Int(Val(rs!TGT_CFC))
    .MSC.TextMatrix(.MSC.Rows - 1, Val(rs!WEK_IDY) * 2) = Int(Val(rs!TGT_QTY))
li:
rs.MoveNext
Next

If rs.State = 1 Then rs.Close
'If .C_SMN_IDY = "All" Then
'    rs.Open "SELECT * FROM PGR WHERE GRP_CAT NOT IN(select PGT_CAT from TGC WHERE SMN_IDY LIKE '0') GROUP BY GRP_CAT", CON, adOpenKeyset, adlockpessimistic
'Else
    rs.Open "SELECT * FROM PGR WHERE GRP_CAT NOT IN(select PGT_CAT from TGC WHERE SMN_IDY LIKE '" & Trim(.smnid) & "') GROUP BY GRP_CAT", CON, adOpenKeyset, adLockPessimistic
'End If
For A = 1 To rs.RecordCount
    .MSC.AddItem ""
    .MSC.TextMatrix(.MSC.Rows - 1, 0) = Trim(rs!GRP_CAT)
    rs.MoveNext
Next
.MousePointer = vbDefault
End With
End Sub

Public Sub MRG_BRD()
If CHECKTB("prd.dbf,tgt.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If
Dim rs As New ADODB.Recordset
With M_FRM_TGT
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRD GROUP BY BRD_NME", CON, adOpenKeyset, adLockPessimistic
.msb.Rows = 2
For A = 1 To rs.RecordCount
    .msb.AddItem ""
    .msb.TextMatrix(.msb.Rows - 1, 0) = Trim(rs!BRD_NME)
    rs.MoveNext
Next
If rs.State = 1 Then rs.Close
'If .SMNID = "All" Then
'    rs.Open "SELECT * FROM TGT,PRD WHERE TGT.PRD_IDY LIKE PRD.PRD_IDY AND SMN_IDY LIKE '0'", CON, adOpenKeyset, adlockpessimistic
'Else
    rs.Open "SELECT * FROM TGT,PRD WHERE TGT.PRD_IDY LIKE PRD.PRD_IDY AND SMN_IDY LIKE '" & Trim(.smnid) & "'", CON, adOpenKeyset, adLockPessimistic
'End If
While Not rs.EOF
    For A = 2 To .msb.Rows - 1
        If Trim(rs!BRD_NME) = Trim(.msb.TextMatrix(A, 0)) Then
            .msb.TextMatrix(A, (Val(rs!WEK_IDY) * 2) - 1) = Int(Val(rs!TGT_CFC)) + Val(.msb.TextMatrix(A, (Val(rs!WEK_IDY) * 2) - 1))
            .msb.TextMatrix(A, Val(rs!WEK_IDY) * 2) = Int(Val(rs!TGT_QTY)) + Val(.msb.TextMatrix(A, (Val(rs!WEK_IDY) * 2)))
            GoTo nxt
        End If
    Next
nxt:
    rs.MoveNext
Wend
.SSTab1.Tab = 1
.MS.Rows = 2
.SSTab1.TabEnabled(0) = False
.SSTab1.TabEnabled(1) = True
.SSTab1.TabEnabled(2) = False
.SSTab1.TabEnabled(3) = False
End With
End Sub

Public Sub MRG_SEG()
If CHECKTB("tgt.dbf,prd.dbf,pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

Dim rs As New ADODB.Recordset
With M_FRM_TGT
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM pgr GROUP BY grp_cls", CON, adOpenKeyset, adLockPessimistic
.MSS.Rows = 2
While Not rs.EOF
    .MSS.AddItem ""
    .MSS.TextMatrix(.MSS.Rows - 1, 0) = Trim(rs!GRP_CLS)
    rs.MoveNext
Wend
If rs.State = 1 Then rs.Close
If .smnid = "All" Then
    rs.Open "SELECT * FROM TGT,PRD,pgr WHERE prd.grp_idy like pgr.grp_idy and TGT.PRD_IDY LIKE PRD.PRD_IDY AND SMN_IDY LIKE '0'", CON, adOpenKeyset, adLockPessimistic
Else
    rs.Open "SELECT * FROM TGT,PRD,pgr WHERE prd.grp_idy like pgr.grp_idy and TGT.PRD_IDY LIKE PRD.PRD_IDY AND SMN_IDY LIKE '" & Trim(.smnid) & "'", CON, adOpenKeyset, adLockPessimistic
End If
While Not rs.EOF
    For A = 2 To .MSS.Rows - 1
        If Trim(rs!GRP_CLS) = Trim(.MSS.TextMatrix(A, 0)) Then
            .MSS.TextMatrix(A, (Val(rs!WEK_IDY) * 2) - 1) = Int(Val(rs!TGT_CFC)) + Val(.MSS.TextMatrix(A, (Val(rs!WEK_IDY) * 2) - 1))
            .MSS.TextMatrix(A, Val(rs!WEK_IDY) * 2) = Int(Val(rs!TGT_QTY)) + Val(.MSS.TextMatrix(A, (Val(rs!WEK_IDY) * 2)))
            GoTo nxt
        End If
    Next
nxt:
    rs.MoveNext
Wend
.SSTab1.Tab = 2
.MS.Rows = 2
.SSTab1.TabEnabled(0) = False
.SSTab1.TabEnabled(2) = True
.SSTab1.TabEnabled(1) = False
.SSTab1.TabEnabled(3) = False
End With
End Sub

Public Sub MRG_CAT()
If CHECKTB("tgs.dbf,pgr.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

Dim rs As New ADODB.Recordset
With M_FRM_TGT
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PGR GROUP BY GRP_CAT", CON, adOpenKeyset, adLockPessimistic
.MSC.Rows = 2
While Not rs.EOF
    .MSC.AddItem ""
    .MSC.TextMatrix(.MSC.Rows - 1, 0) = Trim(rs!GRP_CAT)
    rs.MoveNext
Wend

If rs.State = 1 Then rs.Close
'If .S_SMN_IDY = "All" Then
'    rs.Open "SELECT * FROM TGS,pgr WHERE TGS.PGR_SEG LIKE PGR.GRP_CLS AND SMN_IDY LIKE '0'", CON, adOpenKeyset, adlockpessimistic
'Else
    rs.Open "SELECT * FROM TGS,pgr WHERE TGS.PGR_SEG LIKE PGR.GRP_CLS AND SMN_IDY LIKE '" & Trim(.smnid) & "'", CON, adOpenKeyset, adLockPessimistic
'End If
While Not rs.EOF
    For A = 2 To .MSC.Rows - 1
        If Trim(rs!GRP_CAT) = Trim(.MSC.TextMatrix(A, 0)) Then
            .MSC.TextMatrix(A, (Val(rs!WEK_IDY) * 2) - 1) = Int(Val(rs!TGT_CFC)) + Val(.MSC.TextMatrix(A, (Val(rs!WEK_IDY) * 2) - 1))
            .MSC.TextMatrix(A, Val(rs!WEK_IDY) * 2) = Int(Val(rs!TGT_QTY)) + Val(.MSC.TextMatrix(A, (Val(rs!WEK_IDY) * 2)))
            GoTo nxt
        End If
    Next
nxt:
    rs.MoveNext
Wend
.SSTab1.Tab = 3
.MSS.Rows = 2
.SSTab1.TabEnabled(0) = False
.SSTab1.TabEnabled(1) = False
.SSTab1.TabEnabled(2) = False
.SSTab1.TabEnabled(3) = True
End With
End Sub
