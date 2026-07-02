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
Dim RS As New ADODB.Recordset
RS.Open "select " & fldnames & " from " & tblName, CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    GetValue = Trim(RS(0) & "")
End If
End Function
Public Sub LOGOUT()
cl_logout
End Sub
Public Function amt(amount As String) As String
A = InStr(amount, ".")
If Len(amount) = 1 And A = 1 Then
     amt = "0.00"
     Exit Function
End If
If A > 1 Then
    If InStr(A + 1, amount, ".") > 0 Then
        amt = "0.00"
    Else
        amt = Mid(amount, 1, A) & Mid(amount, A + 1, 2)
    End If
ElseIf Len(Trim(amount)) = 0 Then
   amt = "0.00"
Else
    amt = amount & ".00"
End If
End Function

