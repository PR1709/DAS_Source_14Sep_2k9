Attribute VB_Name = "modSMSSend"
Dim rec As New ADODB.Recordset
'Dim ConnectionString As String

Dim SqlStatement As String

Public Function GetMasterRecords(Index As Integer, lst As ListBox)
    SqlStatement = ""
    Select Case Index
        Case 0
            SqlStatement = " SELECT ACT_NME, ACT_IDY, ACT_MBL FROM ACT WHERE ACT_TYP='SAL'"
        Case 1
            SqlStatement = " SELECT ACT_NME, ACT_IDY, ACT_MBL FROM ACT WHERE ACT_TYP='PUR'"
        Case 2
            SqlStatement = " SELECT SMN_NME, SMN_IDY FROM SMN "
        Case 3
            SqlStatement = " SELECT MAR_NME, MAR_IDY FROM MAR "
        Case 4
            SqlStatement = " SELECT SMN_NME, SMN_IDY FROM SMN "
        Case 5
            SqlStatement = " SELECT DISTINCT ROU_IDY FROM  RDY"
        Case 6
            SqlStatement = " SELECT LOC_NME, LOC_IDY FROM LOC "
        Case 7
            SqlStatement = " SELECT CAT_NME, CAT_IDY FROM CAT "
    End Select
    FillListBox SqlStatement, lst
 End Function


Public Function GetMasterRecordByAcc(Index As Integer, lv As ListView)
    SqlStatement = ""
    Select Case Index
        Case 0
            SqlStatement = " SELECT act_nme as Name, act_idy as ID, ACT_MBL as Mobile FROM ACT WHERE ACT_TYP='SAL'"
        Case 1
            SqlStatement = " SELECT act_nme as Name, act_idy as ID, ACT_MBL as Mobile FROM ACT WHERE ACT_TYP='PUR'"
        Case 2
            SqlStatement = " SELECT SMN_NME as Name, SMN_IDY as ID, SMN_MBL as Mobile FROM SMN "
    End Select
    FillListView lv, SqlStatement
 End Function

Public Function FillListView(lv As ListView, SqlStatement As String, Optional AccountOther As Integer, Optional GetBy As Integer, Optional Value As String)
    lv.ListItems.CLEAR
    lv.ColumnHeaders.CLEAR
    If Trim(SqlStatement) = "" Then Exit Function
    Dim rec As ADODB.Recordset
    Dim li As ListItem
    Dim sql As String
    Dim SqlStatement1 As String
    Dim STR As String
    
'0-Market
'1-Salesman
'2-Route
'3-Location
'4-Channel
'5-Supplier
'6-Retailer

    
    SqlStatement1 = SqlStatement
    If AccountOther = 1 Then
        If GetBy = 0 Then SqlStatement1 = SqlStatement1 & " and act.mar_idy in  (select mar_idy from mar where mar_nme like '" & Trim(Value) & "%' )"
        If GetBy = 1 Then SqlStatement1 = SqlStatement1 & " and act.smn_idy  in (select smn_idy from smn where smn_nme like '" & Trim(Value) & "%' "
        If GetBy = 2 Then SqlStatement1 = SqlStatement1 & " and act.rou_idy like '" & Trim(Value) & "%' "
        If GetBy = 3 Then SqlStatement1 = SqlStatement1 & " and act.loc_idy in (select loc_idy from loc where loc_nme like '" & Trim(Value) & "%' "
        If GetBy = 4 Then SqlStatement1 = SqlStatement1 & " and act.cat_idy in (select cat_idy from cat where cat_nme like '" & Trim(Value) & "%' "
        If GetBy = 5 Then SqlStatement1 = SqlStatement1 & " and act.act_typ = 'PUR' and act.act_nme like '" & Trim(Value) & "%' "
        If GetBy = 6 Then SqlStatement1 = SqlStatement1 & " and act.act_typ = 'SAL' and act.act_nme like '" & Trim(Value) & "%' "
    End If
    Set rec = modGeneral.GetRecords(SqlStatement1)
    For I = 0 To rec.Fields.count - 1
        STR = Trim(rec.Fields(I).name)
        STR = ProperCase(STR)
        'str = UCase(Mid(Trim(rec.Fields(I).name), 1, 1)) & Mid(Trim(rec.Fields(I).name), 2)
        If InStr(1, UCase(rec(I).name), "AMT") > 0 Or InStr(1, UCase(rec(I).name), "AMOUNT") > 0 Then
            lv.ColumnHeaders.Add , , STR, , lvwColumnRight
        ElseIf InStr(1, UCase(rec(I).name), "DATE") > 0 Or InStr(1, UCase(rec(I).name), "DT") > 0 Or InStr(1, UCase(rec(I).name), "ID") > 0 Or InStr(1, UCase(rec(I).name), "DAY") > 0 Or InStr(1, UCase(rec(I).name), "MOBILE") > 0 Then
            If lv.ColumnHeaders.count > 0 Then
                lv.ColumnHeaders.Add , , STR, , lvwColumnCenter
            Else
                lv.ColumnHeaders.Add , , STR, , lvwColumnLeft
            End If
        Else
            lv.ColumnHeaders.Add , , STR, , lvwColumnLeft
        End If
    Next
    While Not rec.EOF
        Set li = lv.ListItems.Add(, , Trim(rec(0)))
        For I = 1 To rec.Fields.count - 1
            If InStr(1, UCase(rec(I).name), "AMT") > 0 Then
                li.ListSubItems.Add , , IIf(IsNull(rec(I)) = True, "", Format(Trim(rec(I)), "#########.00"))
            Else
                li.ListSubItems.Add , , IIf(IsNull(rec(I)) = True, "", Trim(rec(I)))
            End If
        Next
        rec.MoveNext
    Wend
    rec.Close
End Function

Public Function FillListBox(SqlStatement As String, lst As ListBox)
    lst.CLEAR
    Set rec = modGeneral.GetRecords(SqlStatement)
    While Not rec.EOF
        If rec.Fields.count = 1 Then
            lst.AddItem rec(0)
        ElseIf rec.Fields.count = 2 Then
            lst.AddItem rec(0) & "                                                                      | " & rec(1) & "|"
        Else
            lst.AddItem rec(0) & "                                                                      | " & rec(1) & "                                                                      | " & rec(2) & "|"
        End If
        rec.MoveNext
    Wend
    rec.Close
End Function


