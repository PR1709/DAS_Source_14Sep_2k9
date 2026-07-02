Attribute VB_Name = "modInvoice"
Dim rec As New ADODB.Recordset
Dim ConnectionString As String
Public strInvTemplate As String
Dim StrArr() As String

Public Function OpenInvoiceTemplate1(ByRef RTB As RichTextBox) As Long
    On err GoTo ErrHandler
    Dim fso As New FileSystemObject
    Dim f As File
    Dim fileName As String
    Dim ts As TextStream
    Dim I As Integer
    ReDim Preserve StrArr(1)
    fileName = App.Path & "\VATINVOICE.txt"
    RTB.fileName = fileName
    fileName = App.Path & "\SettingDescription.txt"
    Set f = fso.GetFile(fileName)
    Set ts = f.OpenAsTextStream(ForReading)
    I = 1
    While ts.AtEndOfLine = False
        ReDim Preserve StrArr(I)
        StrArr(I - 1) = ts.ReadLine()
        I = I + 1
    Wend
    OpenInvoiceTemplate1 = 0
    Exit Function
ErrHandler:
    OpenInvoiceTemplate1 = err.Number
End Function

Public Function GetRecords(SqlStatement As String) As Recordset
If con.State <> 1 Then
    ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=syam$1234;Persist Security Info=True;User ID=system"
    con.ConnectionString = ConnectionString
    con.Open
End If
rec.Open SqlStatement, con, adOpenStatic, adLockPessimistic
Set GetRecords = rec
End Function

Public Function CreateTextReport(ByRef RTB As RichTextBox, invoiceIds As String)
Dim I As Integer
Dim J As Integer
Dim TOTQTY As Double
Dim VATAMT As Double
Dim DISAMT As Double

Dim STR1 As String
If UBound(StrArr) <= 0 Then
OpenInvoiceTemplate1 RTB
End If
Dim STRSQL As String

If con.State <> 1 Then
    'con.ConnectionString = "Driver=Microsoft Visual FoxPro Driver;SourceType=DBf;SourceDB=D:\MAHAVIR-0809;BackgroundFetch=No;"

    ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=syam$1234;Persist Security Info=True;User ID=system"
    con.ConnectionString = ConnectionString
    con.Open
End If
Dim REC1 As New Recordset
Dim REC2 As New Recordset
STRSQL = "SELECT * FROM INL WHERE TXN_TYP= " & "'SAL'"
    If rec.State = 1 Then rec.Close
    If invoiceIds <> "" Then
        STRSQL = STRSQL & " AND INV_IDY IN (" & invoiceIds & ")"
    End If
        rec.Open STRSQL, con, adOpenStatic, adLockPessimistic
    Dim T As Integer
    
    I = 1
If Not rec.EOF Then
    Do While Not rec.EOF
        If REC1.State = 1 Then REC1.Close
        REC1.Open "SELECT * FROM INV, PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND  INV_IDY = '" & Trim(rec!INV_IDY) & "'", con, adOpenStatic, adLockPessimistic
        J = 1
        T = 0
        TOTQTY = 0
        VATAMT = 0
        DISAMT = 0
        If REC1.EOF = False Then
            While Not REC1.EOF
                DoEvents
                STR1 = I & "-" & Format(J, "0#")
                If I Mod 2 = 0 Then
                    RTB.Text = Replace(RTB.Text, "S" & STR1, SetWord(20, Format(J, "0#"), 1))
                    T = 1
                Else
                    T = 0
                    RTB.Text = Replace(RTB.Text, "S" & STR1, SetWord(3, Format(J, "0#"), 1))
                End If
                
                RTB.Text = Replace(RTB.Text, "PNO" & STR1, SetWord(30, RTrim(REC1!PRD_NME), 1))
                RTB.Text = Replace(RTB.Text, "Q" & STR1, SetWord(9 + T, REC1!PRD_QTY, 0, "######.00"))
                RTB.Text = Replace(RTB.Text, "R" & STR1, SetWord(9 + T, REC1!PRD_RAT, 0, "######.00"))
                RTB.Text = Replace(RTB.Text, "V" & STR1, SetWord(9 + T, REC1!TAX_AMT, 0, "######.00"))
                RTB.Text = Replace(RTB.Text, "D" & STR1, SetWord(9 + T, IIf(IsNull(REC1!DIS_AMT), 0, REC1!DIS_AMT), 0, "######.00"))
                RTB.Text = Replace(RTB.Text, "A" & STR1, SetWord(9 + T, REC1!COL_AMT, 0, "######.00"))
                TOTQTY = TOTQTY + Val(Trim(REC1!PRD_QTY))
                VATAMT = VATAMT + Val(Trim(REC1!TAX_AMT))
                If REC1!DIS_AMT <> Null Then
                DISAMT = DISAMT + Val(Trim(REC1!DIS_AMT))

                End If
                
                J = J + 1
                REC1.MoveNext
            Wend
            If I Mod 2 = 0 Then
                T = 0
            Else
                T = 0
            End If
            RTB.Text = Replace(RTB.Text, "GAMT" & I, SetWord(9 + T, Trim(rec!GRS_AMT), 0, "######.00"))
            RTB.Text = Replace(RTB.Text, "SAMT" & I, SetWord(9 + T, Val(Trim(rec!GRS_AMT)) - DISAMT, 0, "######.00"))
            RTB.Text = Replace(RTB.Text, "VAMT" & I, SetWord(9 + T, CStr(VATAMT), 0, "######.00"))
            RTB.Text = Replace(RTB.Text, "DAMT" & I, SetWord(9 + T, CStr(DISAMT), 0, "######.00"))
            RTB.Text = Replace(RTB.Text, "TOTQTY" & I, SetWord(9 + T, Val(TOTQTY), 0, "######.00"))
            RTB.Text = Replace(RTB.Text, "TOTSAL" & I, SetWord(9 + T, CStr(Trim(rec!GRS_AMT)), 0, "######.00"))
            RTB.Text = Replace(RTB.Text, "TOTTAX" & I, SetWord(9 + T, CStr(VATAMT), 0, "######.00"))
            RTB.Text = Replace(RTB.Text, "TOTAMT" & I, SetWord(9 + T, Trim(rec!NET_AMT), 0, "######.00"))
            RTB.Text = Replace(RTB.Text, "TOTAMT" & I, SetWord(9 + T, " ", 0, "######.00"))
            RTB.Text = Replace(RTB.Text, "RNDOFF" & I, SetWord(9 + T, " ", 0, "######.00"))
            RTB.Text = Replace(RTB.Text, "NETAMT" & I, SetWord(9 + T, " ", 0, "######.00"))
            For K = J - 1 To 11
                DoEvents
                STR1 = I & "-" & Format(K, "0#")
                RTB.Text = Replace(RTB.Text, "S" & STR1, SetWord(3, " ", 0))
                RTB.Text = Replace(RTB.Text, "PNO" & STR1, SetWord(66, " ", 0))
                RTB.Text = Replace(RTB.Text, "Q" & STR1, SetWord(9, " ", 0))
                RTB.Text = Replace(RTB.Text, "R" & STR1, SetWord(9, " ", 0))
                RTB.Text = Replace(RTB.Text, "V" & STR1, SetWord(9, " ", 0))
                RTB.Text = Replace(RTB.Text, "D" & STR1, SetWord(9, " ", 0))
                RTB.Text = Replace(RTB.Text, "A" & STR1, SetWord(9, " ", 0))
            Next
            I = I + 1
        End If
        If I = 5 Then
            rec.Close
            REC1.Close
            Exit Do
        End If
        rec.MoveNext
    Loop
End If
For I = I To 4
        For J = 1 To 10
            DoEvents
            STR1 = I & "-" & Format(J, "0#")
            RTB.Text = Replace(RTB.Text, "S" & STR1, " ")
            RTB.Text = Replace(RTB.Text, "PNO" & STR1, " ")
            RTB.Text = Replace(RTB.Text, "Q" & STR1, " ")
            RTB.Text = Replace(RTB.Text, "R" & STR1, " ")
            RTB.Text = Replace(RTB.Text, "V" & STR1, " ")
            RTB.Text = Replace(RTB.Text, "D" & STR1, " ")
            RTB.Text = Replace(RTB.Text, "A" & STR1, " ")
        Next
        RTB.Text = Replace(RTB.Text, "GAMT" & I, "")
        RTB.Text = Replace(RTB.Text, "SAMT" & I, " ")
        RTB.Text = Replace(RTB.Text, "VAMT" & I, " ")
        RTB.Text = Replace(RTB.Text, "DAMT" & I, " ")
        RTB.Text = Replace(RTB.Text, "TOTQTY" & I, " ")
        RTB.Text = Replace(RTB.Text, "TOTSAL" & I, " ")
        RTB.Text = Replace(RTB.Text, "TOTTAX" & I, " ")
        RTB.Text = Replace(RTB.Text, "TOTAMT" & I, " ")
Next
End Function

Public Function SetOthers(ByRef RTB As RichTextBox)
OpenInvoiceTemplate1 RTB
Dim J As Integer
Dim fname As String
Dim fvalue As String
Dim actualvalue As Integer

J = 1
    For J = 1 To 2
        For I = 0 To UBound(StrArr) - 1
        fname = Split(StrArr(I), "|")(0) & J
        fvalue = Split(StrArr(I), "|")(1)
        actualvalue = CInt(Split(StrArr(I), "|")(2))
            RTB.Text = Replace(RTB.Text, fname, SetWord(actualvalue, fvalue, 1))
        Next
    Next
End Function


Private Function SetWord(ActualLen As Integer, value As String, LeftRight As Integer, Optional formatString As String) As String
    Dim LEN1 As Integer
    Dim I As Integer
    If Len(formatString) > 0 Then
        value = CStr(Format(value, formatString))
    End If
    LEN1 = Len(value)
    For I = LEN1 To ActualLen
'        If LeftRight = 0 Then
        value = Chr(32) & value
'        Else
'        value = value & Chr(32)
'        End If
    Next
    SetWord = value
End Function
