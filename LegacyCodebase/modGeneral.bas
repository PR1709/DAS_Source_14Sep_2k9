Attribute VB_Name = "modGeneral"


Public Function GetRecords(SqlStatement As String) As ADODB.Recordset
On Error GoTo ErrHandler
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    If con.State = 1 Then con.Close
        con.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & "\EFF.DBC;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    rec.CursorLocation = adUseClient
    rec.Open SqlStatement, con, adOpenKeyset
    Set GetRecords = rec
    Exit Function
ErrHandler:
    Set GetRecords = Nothing
End Function

Public Function ProperCase(str As String) As String
str = UCase(Mid(str, 1, 1)) & Mid(str, 2)
For I = 2 To Len(str)
    If Asc(Mid(str, I, 1)) = vbKeySpace Then
        str = Mid(str, 1, I) & UCase(Mid(str, I + 1, 1)) & Mid(str, I + 2)
    End If
Next
     
    ProperCase = str
End Function
