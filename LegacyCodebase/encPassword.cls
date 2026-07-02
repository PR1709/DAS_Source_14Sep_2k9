VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "encPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function encPwd(pw As String)
cfm = Format(Now, "ddMMyyyyhhmmss")
dtm = StrReverse(cfm)
num = 0
For I = 1 To Len(Str(dtm))
    num = num + (Val(Mid(dtm, I, 1)) + I)
Next
For I = 1 To Len(Trim(pw))
    pwd = pwd & Abs(Asc(Mid(Trim(pw), I, 1)) - num) & "-"
Next
pwd = pwd & Chr(num)
encPwd = pwd
End Function
