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

Public Sub LOGOUT()
cl_logout
End Sub
