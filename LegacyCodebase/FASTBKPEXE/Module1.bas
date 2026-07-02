Attribute VB_Name = "Module1"
Public Declare Function cl_login Lib "CL32.DLL" (ByVal dKey As String) As Integer
Public Declare Function cl_logout Lib "CL32.DLL" () As Integer
Public Declare Function cl_lock_ok Lib "CL32.DLL" () As Integer
Public Sub LOGIN()
If cl_login("067670508880070277000726") <> 1 Then
    MsgBox "DAS Version 4.0 License Dongle Missing! Check Device Connection to Parallel Port - USB!", vbCritical, "DAS Version 4.0"
    End
End If
End Sub
Public Function GetValue(fldnames As String, tblName As String) As String
'If CHECKTB(tblname & ".dbf") = 1 Then
'    MsgBox UCase(tblname) & " Missing "
'    Exit Function
'End If
Dim CON As New ADODB.Connection
CON.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CurDir & "\EFF.DBC;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
Dim rs As New ADODB.Recordset
rs.Open "select " & fldnames & " from " & tblName, CON, adOpenKeyset, adLockOptimistic
If rs.RecordCount > 0 Then
    GetValue = Trim(rs(0) & "")
End If
End Function
Public Sub LOGOUT()
cl_logout
End Sub
