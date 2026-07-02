VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "decPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function decPwd(pw As String) As String
st = Right(Trim(pw), 1)
ast = Asc(st)
st = Mid(Trim(pw), 1, Len(Trim(pw)) - 1)
arr = Split(st, "-", , vbTextCompare)
pwd = ""
For i = 0 To UBound(arr) - 1
    pwd = pwd & Chr(ast - arr(i))
Next
decPwd = pwd
End Function
