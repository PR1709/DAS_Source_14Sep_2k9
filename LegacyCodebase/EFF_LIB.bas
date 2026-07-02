Attribute VB_Name = "EFF_LIB"
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function cl_login Lib "CL32.DLL" (ByVal dKey As String) As Integer
Public Declare Function cl_logout Lib "CL32.DLL" () As Integer
Public Declare Function cl_lock_ok Lib "CL32.DLL" () As Integer
Public SEL_CMP As String
Public DEL_REC_IDY As String
Public CURDIR As String
Public SAL_QRY As String
Public CON As New ADODB.Connection
Public PNV As New ADODB.Connection
Public PNL As New ADODB.Connection
Public RRI As New ADODB.Connection
Public INVOICE_DATE As Date
Public LSS_TYP As String
Public FRM_INV_DAT As Date
Public TO_INV_DAT As Date
Public to_dat As Date
Public FRM_INV As String
Public END_INV As String
Public ROU_IDY As String
Public INV_TYP As String
Public USERNAME
Public F2, F3, F4, F5, F6, F7, F8, F9 As String
Public Type fontdetails
name As String
size As Integer
bold As Boolean
italic As Boolean
End Type
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public nid As NOTIFYICONDATA
Public FRM_DAT As Date
Public TO_DATE As Date
Public sal_act As String
Public CN As New ADODB.Connection
Public CN1 As New ADODB.Connection
Public cn2 As New ADODB.Connection
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public db As Database
Public RS As New ADODB.Recordset
Public tb As TableDef
Public DB1 As Database
Public RS1 As New ADODB.Recordset
Public tb1 As TableDef
Public dbs As Database

Public Function exitform(FRM As Form)
'FRM.Command1.SetFocus
End Function

Public Sub clrctr(FRM As Form)
For A = 1 To FRM.Count - 1
If TypeOf FRM.Controls(A) Is TextBox Then FRM.Controls(A).Text = ""
If TypeOf FRM.Controls(A) Is MaskEdBox Then FRM.Controls(A).Text = ""
Next
End Sub

Public Function convert(num As String) As String
If Len(Trim(num)) >= 10 Then
    convert = Round(num, 2)
Else
    num = Round(Val(num), 2)
    If InStr(num, ".") > 0 Then
        r = Mid(num, InStr(num, ".") + 1, Len(num) - InStr(num, "."))
        If Trim(Len(r)) = 1 Then
            convert = Mid(num, 1, InStr(num, ".") - 1) & "." & r & "0"
        ElseIf Trim(Len(r)) = 2 Then
            convert = num
        End If
    Else
        If Trim(Len(num)) <> 0 Then
            convert = num & ".00"
        Else
            convert = "0.00"
        End If
        End If
End If
End Function

Public Function check(tbnme As String, fldnme As String, fldvalue As String) As Boolean
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from " & tbnme & " where " & fldnme & " like '" & Trim(fldvalue) & "'", CON, adOpenStatic
If RS.RecordCount > 0 Then
check = True
Else
check = False
End If
If RS.State = 1 Then RS.Close
End Function

Public Function querycheck(tbnme As String, fldnme As String, fldvalue As String, qry As String) As Boolean
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from " & tbnme & " where " & fldnme & " like '" & Trim(fldvalue) & "' " & Trim(qry), CON, adOpenStatic
If RS.RecordCount > 0 Then
querycheck = True
Else
querycheck = False
End If
If RS.State = 1 Then RS.Close
End Function

Public Function Check1(tbnme As String, fldnme As String, fldvalue As String, typ As String) As Boolean
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select * from " & tbnme & " where " & fldnme & " like '" & Trim(fldvalue) & "' ", CON, adOpenStatic
If RS.RecordCount > 0 Then
Check1 = True
Else
Check1 = False
End If
If RS.State = 1 Then RS.Close
End Function

Public Function increment(fldval As String) As String
Dim ict, icr As Integer, icrt As String, rec As New ADODB.Recordset
If rec.State = 1 Then rec.Close
rec.Open "select " & fldval & " from idy", CON, adOpenKeyset, adLockOptimistic
If rec.RecordCount = 0 Then
rec.AddNew
rec.Update
End If
rec(0) = Val(rec(0) & "") + 1
ict = Val(rec(0) & "")
rec.Update
If rec.State = 1 Then rec.Close
rec.Open "select pre_" & Mid(fldval, 5, 3) & " from pre", CON, adOpenKeyset, adLockOptimistic
If Len(Trim(rec(0) & "") & ict) = 10 Then
    increment = Trim(rec(0) & "") & ict
Else
    icrt = Trim(rec(0) & "")
    Do While Len(icrt) <= 10 - Len(STR(ict))
    icrt = icrt & "0"
    Loop
End If
increment = icrt & ict
End Function

Public Function T7increment(fldval As String) As String
Dim ict, icr As Integer, icrt As String, rec As New ADODB.Recordset
If rec.State = 1 Then rec.Close
rec.Open "select " & fldval & " from idy", CON, adOpenKeyset, adLockOptimistic
If rec.RecordCount = 0 Then
rec.AddNew
rec.Update
End If
rec(0) = Val(rec(0) & "") + 1
ict = Val(rec(0) & "")
rec.Update
Do While Len(icrt) <= 7 - Len(STR(ict))
icrt = icrt & "0"
Loop
T7increment = icrt & ict
End Function

Public Function Tincrement(fldval As String) As String
'transactions increment
If Len(fldval) < 7 Then
icrt = ""
Do While Len(icrt) < (7 - Len(Trim(fldval)))
    icrt = icrt & "0"
Loop
Tincrement = Trim(icrt & fldval)
Else
Tincrement = fldval
End If
End Function

Public Function Sincrement(fldval As String) As String
'retreving transactions
Do While Len(icrt) < 10 - Len(fldval)
    icrt = icrt & "0"
Loop
Sincrement = Trim(icrt & fldval)
End Function

Public Function labdisp(fldval As String) As String
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select " & fldval & " from idy", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount = 0 Then
RS.AddNew "IDY_OPS", 0
RS.Update
End If
If Val(RS(0) & "") = 0 Then
labdisp = "0000000"
ElseIf Val(RS(0)) > 0 And Val(RS(0)) <= 9 Then
labdisp = "000000" & Val(RS(0))
ElseIf Val(RS(0)) >= 10 And Val(RS(0)) <= 99 Then
labdisp = "00000" & Val(RS(0))
ElseIf Val(RS(0)) >= 100 And Val(RS(0)) <= 999 Then
labdisp = "0000" & Val(RS(0))
ElseIf Val(RS(0)) >= 1000 And Val(RS(0)) <= 9999 Then
labdisp = "000" & Val(RS(0))
ElseIf Val(RS(0)) >= 10000 And Val(RS(0)) <= 99999 Then
labdisp = "00" & Val(RS(0))
ElseIf Val(RS(0)) >= 100000 And Val(RS(0)) <= 999999 Then
labdisp = "0" & Val(RS(0))
ElseIf Val(RS(0)) >= 1000000 And Val(RS(0)) <= 9999999 Then
labdisp = Val(RS(0))
End If
End Function

Public Function discount(dis As String) As String
A = InStr(dis, ".")
b = InStr(dis, "-")
If InStr(A + 1, dis, ".") > 0 Or (Len(Trim(dis)) = 1 And A = 1) Then
    discount = "0.00"
    Exit Function
End If
If b = 1 And Len(dis) = 1 Then
    discount = "0.00"
    Exit Function
ElseIf b > 1 And Len(dis) >= 2 Then
    discount = "0.00"
    Exit Function
ElseIf b = 1 And Len(dis) >= 5 And A = 0 Then
    discount = Mid(dis, 1, 3) & "." & Mid(dis, 4, 2)
    Exit Function
ElseIf b = 1 And Len(dis) >= 5 And A = 4 Then
    discount = dis
    Exit Function
ElseIf b = 1 And Len(dis) = 3 And A = 0 Then
    discount = dis & ".00"
    Exit Function
ElseIf b = 1 And Len(dis) >= 3 And A = 2 Then
    discount = Mid(dis, 1, 1) & "0" & Mid(dis, 2, 3)
    Exit Function
    
ElseIf b = 1 And Len(dis) = 2 And A = 2 Then
    discount = "0.00"
    Exit Function
ElseIf b = 1 And Len(dis) = 5 And A = 5 Then
    discount = Mid(dis, 1, 3) & "." & Mid(dis, 4, 1) & "0"
    Exit Function
ElseIf b = 1 And Len(dis) > 5 And A = 5 Then
    discount = Mid(dis, 1, 3) & "." & Mid(dis, 4, 1) & Mid(dis, 6, 1)
    Exit Function
ElseIf b = 1 And Len(dis) > 5 And A = 6 Then
    discount = Mid(dis, 1, 3) & "." & Mid(dis, 4, 2)
    Exit Function
End If
If Len(dis) >= 4 And A = 0 Then
    discount = Mid(dis, 1, 2) & "." & Mid(dis, 3, 2)
ElseIf Len(dis) > 4 And A = 4 Then
    discount = Mid(dis, 1, 2) & "." & Mid(dis, 3, 1) & Mid(dis, 5, 1)
ElseIf Len(dis) = 4 And A = 4 Then
    discount = Mid(dis, 1, 2) & "." & Mid(dis, 3, 1) & "0"
ElseIf Len(dis) >= 4 And A = 2 Then
    discount = Mid(dis, 1, 2) & Mid(dis, 3, 2)
ElseIf Len(dis) = 4 And A = 3 Then
    discount = Mid(dis, 1, 3) & Mid(dis, 4, 1) & "0"
ElseIf Len(dis) = 6 And A = 3 Then
    discount = Mid(dis, 1, 3) & Mid(dis, 4, 2)
ElseIf Len(dis) >= 4 And (A = 5 Or A = 6) Then
    discount = Mid(dis, 1, 2) & "." & Mid(dis, 3, 2)
ElseIf Mid(dis, 1, 1) = "." And Mid(dis, 2, 1) = "." Then
    discount = "0.00"
ElseIf Len(dis) >= 3 And A = 1 Then
    discount = "0" & Mid(dis, 1, 3)
ElseIf Len(dis) = 3 And A = 0 Then
    discount = Mid(dis, 1, 2) & "." & Mid(dis, 3, 1) & "0"
ElseIf Len(dis) = 3 And A = 2 Then
    discount = Mid(dis, 1, 3) & "0"
ElseIf Len(dis) = 3 And A = 3 Then
    discount = Mid(dis, 1, 3) & "00"
ElseIf Len(dis) = 2 And A = 1 Then
    discount = "0" & Mid(dis, 1, 2) & "0"
ElseIf Len(dis) = 2 And A = 2 Then
    discount = Mid(dis, 1, 2) & "00"
ElseIf (Len(dis) = 1 Or Len(dis) = 2) And A = 0 Then
    discount = dis & ".00"
ElseIf Len(dis) = 0 Then
    discount = "0.00"
Else
    discount = dis
End If
End Function

Public Function AMT(amount As String) As String
A = InStr(amount, ".")
If Len(amount) = 1 And A = 1 Then
     AMT = "0.00"
     Exit Function
End If
If A > 1 Then
    If InStr(A + 1, amount, ".") > 0 Then
        AMT = "0.00"
    Else
        AMT = Mid(amount, 1, A) & Mid(amount, A + 1, 2)
    End If
ElseIf A = 1 And Len(amount) = 2 Then
    AMT = "0" & amount & "0"
ElseIf A = 1 And Len(amount) >= 3 Then
    AMT = "0" & Mid(amount, 1, 3)
ElseIf Len(Trim(amount)) > 7 Then
    AMT = Mid(amount, 1, 6) & "." & Mid(amount, 7, 2)
ElseIf Len(Trim(amount)) = 7 Then
    AMT = Mid(amount, 1, 6) & "." & Mid(amount, 7, 1) & "0"
ElseIf Len(Trim(amount)) = 0 Then
   AMT = "0.00"
ElseIf Len(Trim(amount)) < 7 Then
    AMT = amount & ".00"
End If
End Function

Public Function QTY(amount As String) As String
b = InStr(amount, ".")
If b <> 0 Then
    M = Mid(amount, b + 1, Len(amount) - (b))
    If Len(M) = 1 Then
        QTY = amount & "00"
    ElseIf Len(M) = 2 Then
        QTY = amount & "0"
    ElseIf Len(M) = 0 Then
        QTY = amount & "000"
    ElseIf Len(M) = 3 Then
        QTY = amount
    Else
        QTY = "0.000"
    End If
ElseIf Len(Trim(amount)) > 5 Then
    M = Mid(amount, 1, 5)
    DEC = Mid(amount, 6, Len(amount) - 5)

    If Len(DEC) = 1 Then
        QTY = M & "." & DEC & "00"
    ElseIf Len(DEC) = 2 Then
        QTY = M & "." & DEC & "0"
    ElseIf Len(DEC) = 0 Then
        QTY = M & ".000"
    ElseIf Len(DEC) = 3 Then
        QTY = M & "." & DEC
    ElseIf Len(DEC) = 4 Then
        QTY = M & "." & Mid(Trim(DEC), 1, 3)
    End If
Else
    QTY = amount & ".000"
End If
End Function

Public Function FAC(amount As String) As String
A = InStr(amount, ".")
If Len(amount) = 1 And A = 1 Then
     FAC = "0.000"
     Exit Function
End If
If A > 1 Then
    If InStr(A + 1, amount, ".") > 0 Then
        FAC = "0.000"
    Else
        FAC = Mid(amount, 1, A) & Mid(amount, A + 1, 3)
                    End If
ElseIf A = 1 And Len(amount) = 2 Then
    FAC = "0" & amount & "00"
ElseIf A = 1 And Len(amount) = 3 Then
    FAC = "0" & Mid(amount, 1, 3) & "0"
ElseIf A = 1 And Len(amount) >= 4 Then
    FAC = "0" & Mid(amount, 1, 4)
ElseIf Len(Trim(amount)) > 6 Then
    FAC = Mid(amount, 1, 5) & "." & Mid(amount, 6, 3)
ElseIf Len(Trim(amount)) = 6 Then
    FAC = Mid(amount, 1, 5) & "." & Mid(amount, 6, 1) & "00"
ElseIf Len(Trim(amount)) = 0 Then
   FAC = "0.000"
ElseIf Len(Trim(amount)) <= 5 Then
    FAC = amount & ".000"
End If
End Function

Public Function FUNKEY(KEY As Integer) As String
If KEY = 113 Then
    FUNKEY = F2
ElseIf KEY = 114 Then
    FUNKEY = F3
ElseIf KEY = 115 Then
    FUNKEY = F4
ElseIf KEY = 116 Then
    FUNKEY = F5
ElseIf KEY = 117 Then
    FUNKEY = F6
ElseIf KEY = 118 Then
    FUNKEY = F7
ElseIf KEY = 119 Then
    FUNKEY = F8
End If
End Function

Public Function mincrement(nme As String, fldval As String) As String
Do While Len(Trim(icrt)) <= 4 - Len(Trim(fldval))
icrt = icrt & "0"
Loop
mincrement = icrt & fldval
End Function

Public Sub checkdongle()
Dim CL_INT As Integer
CL_INT = cl_lock_ok
If CL_INT <> 1 Then
   MsgBox "Security Device Missing! Check Device Connection to Parallel Port!", vbCritical, "Effmcg"
   End
End If
End Sub

Public Sub LOGIN()
If cl_login("067670508880070277000726") <> 1 Then
MsgBox "Security Device Missing! Check Device Connection to Parallel Port!", vbCritical, "Effmcg"
End
End If
End Sub

Public Sub LOGOUT()
Dim RSlog As New ADODB.Recordset
RSlog.Open "select * from log", CON, adOpenKeyset, adLockOptimistic
If RSlog.EOF = False Then
   RSlog.MoveLast
   RSlog!lot_dat = Date
   RSlog!lot_tim = Format(Time, "HH:mm:SS")
   RSlog.Update
Else
End If
cl_logout
End Sub

Public Sub SETYEAR(FRM As Form)
On Error GoTo Label
Dim RS As New ADODB.Recordset
RS.Open "SELECT * FROM YER WHERE ACT_PAS LIKE 'Y'", CON, adOpenKeyset
For A = 0 To FRM.Count - 1
    If TypeOf FRM.Controls(A) Is DTPicker Then
        FRM.Controls(A).MaxDate = RS("TO_DAT")
        FRM.Controls(A).MinDate = RS("FRM_DAT")
        FRM.Controls(A).Value = Date
        CON.Execute "UPDATE DIR SET DIR_ALS = '" & Format(RS("FRM_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
        CON.Execute "UPDATE DIR SET DIR_MSG = '" & Format(RS("TO_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
    End If
Next
Exit Sub
Label:
MsgBox "Invalid Accounting Period, Contact Effmcg Support!", vbCritical
End
End Sub

Public Sub SETYEARNEW(FRM As Form)
On Error GoTo Label
Dim RS As New ADODB.Recordset
RS.Open "SELECT * FROM YER WHERE ACT_PAS LIKE 'Y'", CON, adOpenKeyset
For A = 0 To FRM.Count - 1
    If TypeOf FRM.Controls(A) Is DTPicker Then
        FRM.Controls(A).MaxDate = RS("FRM_DAT")
        FRM.Controls(A).MinDate = RS("FRM_DAT")
        FRM.Controls(A).Value = RS("FRM_DAT")
        CON.Execute "UPDATE DIR SET DIR_ALS = '" & Format(RS("FRM_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
        CON.Execute "UPDATE DIR SET DIR_MSG = '" & Format(RS("TO_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
    End If
Next
Exit Sub
Label:
MsgBox "Invalid Accounting Period, Contact Effmcg Support!", vbCritical
End
End Sub

Public Function fragment(AMT As String) As String
If Len(Trim(AMT)) = 10 Then
    fragment = Mid(AMT, 1, 2) & "," & Mid(AMT, 3, 2) & "," & Mid(AMT, Len(AMT) - 6 + 1, 6)
ElseIf Len(Trim(AMT)) = 9 Then
    fragment = Mid(AMT, 1, 1) & "," & Mid(AMT, 2, 2) & "," & Mid(AMT, Len(AMT) - 6 + 1, 6)
ElseIf Len(Trim(AMT)) = 8 Then
    fragment = Mid(AMT, 1, 2) & "," & Mid(AMT, Len(AMT) - 6 + 1, 6)
ElseIf Len(Trim(AMT)) = 7 Then
    fragment = Mid(AMT, 1, 1) & "," & Mid(AMT, Len(AMT) - 6 + 1, 6)
Else
    fragment = AMT
End If
End Function

Public Function CONQTY(ByVal EQTY As Double, prdid As String) As Double
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select sub_unt,cur_lev from prd where prd_idy like '" & Trim(prdid) & "'", CON, adOpenKeyset, adLockOptimistic
Ent_qty = MTS(Val(EQTY), RS(0))
CUR_QTY = MTS(Val(RS(1)), RS(0))
CONQTY = STM(Val(CUR_QTY) - Val(Ent_qty), RS(0))
End Function

Public Function CONPRDQTY(ByVal CURQTY As String, ByVal EQTY As String, prdid As String) As Long
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
EQTY = QTY(EQTY)
CURQTY = QTY(CURQTY)
RS.Open "select sub_unt from prd where prd_idy like '" & prdid & "'", CON, adOpenKeyset, adLockOptimistic
If Val(RS(0)) > 0 Then
    A = InStr(EQTY, ".")
    b = InStr(CURQTY, ".")
    If b > 0 Then
        CDECNUM = Mid(CURQTY, b + 1, 3)
        NUM2 = Mid(CURQTY, 1, b - 1)
        TOT1 = Val(RS(0) * NUM2) + Val(CDECNUM)
    End If
    
    If A > 0 Then
    EDECNUM = Mid(EQTY, A + 1, 3)
    NUM1 = Mid(EQTY, 1, A - 1)
    TOT2 = Val(RS(0) * Val(NUM1)) + Val(EDECNUM)
    End If
    CONPRDQTY = TOT1 - TOT2
Else
    CONPRDQTY = QTY(Round(CURQTY - EQTY, 3))
End If
End Function

Public Function ADDQTY(ByVal EQTY As Double, prdid As String) As Double
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select sub_unt,cur_lev from prd where PRD_idy like '" & Trim(prdid) & "'", CON, adOpenKeyset, adLockOptimistic
Ent_qty = MTS(Val(EQTY), RS(0))
CUR_QTY = MTS(Val(RS(1)), RS(0))
ADDQTY = STM(Val(Ent_qty) + Val(CUR_QTY), RS(0))
End Function

Public Function ADDBRDQTY(ByVal EQTY As Double, BRDID As String) As Double
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select sub_unt,prd_qty from prd,brd where prd.prd_idy like brd.PRD_idy and BRD.bat_idy like '" & Trim(BRDID) & "'", CON, adOpenKeyset, adLockOptimistic
Ent_qty = MTS(Val(EQTY), RS(0))
CUR_QTY = MTS(Val(RS(1) & ""), RS(0))
ADDBRDQTY = STM(Val(Ent_qty) + Val(CUR_QTY), RS(0))
End Function
Public Function CONBRDQTY(ByVal EQTY As Double, BRDID As String) As Double
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
RS.Open "select sub_unt,prd_qty from prd,brd where prd.prd_idy like brd.PRD_idy and bat_idy like '" & Trim(BRDID) & "'", CON, adOpenKeyset, adLockOptimistic
Ent_qty = MTS(Val(EQTY), RS(0))
CUR_QTY = MTS(Val(RS(1)), RS(0))
CONBRDQTY = STM(Val(CUR_QTY) - Val(Ent_qty), RS(0))
End Function

Public Function VANADDQTY(EQTY As String, prdid As String, INVID As String) As String
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If RS.State = 1 Then RS.Close
If RS1.State = 1 Then RS1.Close
EQTY = QTY(EQTY)

RS.Open "select sub_unt from prd where prd_idy like '" & Trim(prdid) & "'", CON, adOpenKeyset, adLockOptimistic
RS1.Open "select RET_QTY from inv where prd_idy like '" & Trim(prdid) & "' and txn_typ like 'RSI' AND inv_idy like '" & Trim(INVID) & "'", CON, adOpenKeyset, adLockOptimistic
If Val(RS(0)) > 0 Then
    A = InStr(QTY(EQTY), ".")
    b = InStr(QTY(RS1(0) & ""), ".")
    
    If A > 0 Then
    EDECNUM = Mid(EQTY, A + 1, 3)
    NUM1 = Mid(EQTY, 1, A - 1)
    TOT1 = Val(Val(RS(0) & "") * Val(NUM1)) + Val(EDECNUM)
    End If
    If b > 0 Then
    CDECNUM = Mid(QTY(RS1(0) & ""), b + 1, 3)
    NUM2 = Mid(QTY(RS1(0) & ""), 1, b - 1)
    TOT2 = Val(Val(RS(0)) * Val(NUM2)) + Val(CDECNUM)
    End If
    'tot2 = Val(rs(0) * num2) + Round(Val(cdecnum), 3)
    tot = TOT1 + TOT2
    r = Round(Val(tot) Mod Val(RS(0)), 3)       'remainder
    V = Val(tot) \ Val(RS(0))   'quotient
    If Len(Trim(r)) = 1 Then
    VANADDQTY = V & ".00" & r
    ElseIf Len(Trim(r)) = 2 Then
    VANADDQTY = V & ".0" & r
    ElseIf Len(Trim(r)) = 3 Then
    VANADDQTY = V & "." & r
    End If
Else
    VANADDQTY = EQTY
End If
VANADDQTY = QTY(VANADDQTY)
End Function

Public Function VANCONQTY(EQTY As String, prdid As String, INVID As String) As String
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If RS.State = 1 Then RS.Close
If RS1.State = 1 Then RS1.Close
EQTY = QTY(EQTY)

RS.Open "select sub_unt from prd where prd_idy like '" & Trim(prdid) & "'", CON, adOpenKeyset, adLockOptimistic
RS1.Open "select RET_qty from inv where prd_idy like '" & Trim(prdid) & "' and txn_typ like 'RSI' AND inv_idy like '" & Trim(INVID) & "'", CON, adOpenKeyset, adLockOptimistic
If Val(RS(0)) > 0 Then
    A = InStr(QTY(EQTY), ".")
    b = InStr(QTY(RS1(0) & ""), ".")
    
    If A > 0 Then
    EDECNUM = Mid(EQTY, A + 1, 3)
    NUM1 = Mid(EQTY, 1, A - 1)
    TOT1 = Val(RS(0) * NUM1) + Val(EDECNUM)
    End If
       
    If b > 0 Then
    CDECNUM = Mid(QTY(RS1(0) & ""), b + 1, 3)
    NUM2 = Mid(QTY(RS1(0) & ""), 1, b - 1)
    TOT2 = (Val(RS(0) & "") * Val(NUM2)) + Val(CDECNUM)
    End If
    'tot2 = Val(rs(0) * num2) + Round(Val(cdecnum), 3)
    tot = TOT1 - TOT2
    
    r = Round(Val(tot) Mod Val(RS(0)), 3)       'remainder
    V = Val(tot) \ Val(RS(0))   'quotient
    If Len(Trim(r)) = 1 Then
        VANCONQTY = V & ".00" & r
    ElseIf Len(Trim(r)) = 2 Then
        VANCONQTY = V & ".0" & r
    ElseIf Len(Trim(r)) = 3 Then
        VANCONQTY = V & "." & r
    End If
Else
    VANCONQTY = EQTY
End If
    VANCONQTY = QTY(VANCONQTY)
End Function

Public Function RATCAL(UNT As String, RAT As String, SUBUNT As String) As String
UNT = QTY(UNT)
A = InStr(QTY(UNT), ".")
If A > 0 Then
    EDECNUM = Mid(UNT, A + 1, 3)
    NUM1 = Mid(UNT, 1, A - 1)
End If
c = (Val(EDECNUM) * Val(RAT))
If c <> 0 Then
V = c / Val(SUBUNT)
Else
V = 0
End If
RATCAL = convert(AMT((Val(NUM1) * Val(RAT)) + Val(V)))
End Function

Public Function WGTCAL(UNT As String, CONFAC As String, SUBUNT As String) As String
UNT = QTY(UNT)
A = InStr(QTY(UNT), ".")
If A > 0 Then
    EDECNUM = Mid(UNT, A + 1, 3)
    NUM1 = Mid(UNT, 1, A - 1)
End If
WGTCAL = QTY((Val(NUM1) * Val(CONFAC)) + Round(((Val(EDECNUM) * Val(CONFAC)) / Val(SUBUNT)), 3))
End Function

Public Function UNTCAL(UNT As String, prdid As String) As String
Dim RS As New ADODB.Recordset
RS.Open "SELECT * FROM PRD WHERE PRD_IDY LIKE '" & Trim(prdid) & "'", CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
A = Int(Val(UNT) / Val(RS("CON_FAC") & ""))
c = Round((Val(RS("CON_FAC") * 1000) / Val(RS("SUB_UNT") * 1000)), 3)
b = Val(UNT) * 1000 Mod c * 1000
If b <> 0 Then MsgBox "Conversion Factor and Sub Units Mismatch!", vbExclamation
b = Val(UNT) * 1000 Mod Val(RS("CON_FAC") & "") * 1000
b = b / (Val(RS("CON_FAC") & "") / Val(RS("SUB_UNT") & "") * 1000)
b = b / 1000
UNTCAL = QTY(Val(A) + Val(b))
Else
UNTCAL = 0
End If
End Function

Public Sub TRAIL_BALANCE(TO_DAT1 As Date)
'If MsgBox("Update Accounts", vbYesNo + vbQuestion) = vbYes Then
to_dat = TO_DAT1
Unload M_FRM_ASD
M_FRM_APD.Tag = "TB"
M_FRM_APD.Show 1
'    Dim RS As New ADODB.Recordset
'    Dim RS1 As New ADODB.Recordset
'    If RS.State = 1 Then RS.Close
'    If RS1.State = 1 Then RS1.Close
'    con.Execute "update opa set opa_dbt=0"
'    con.Execute "update opa set opa_crt=0"
'    RS.Open "select * from vch where txn_dat <ctod('" & Format(TO_DAT, "MM/DD/YYYY") & "')", con, adOpenKeyset, adLockOptimistic
'    If RS.RecordCount > 0 Then
'    For A = 1 To RS.RecordCount
'        If RS1.State = 1 Then RS1.Close
'        RS1.Open "SELECT * FROM opa where act_idy like '" & Trim(RS("dbt_idy")) & "'", con, adOpenKeyset, adLockOptimistic
'        If RS1.RecordCount > 0 Then
'            RS1("opa_dbt") = Val(RS1("opa_dbt")) + Val(RS("txn_amt") & "")
'            RS1.Update
'        Else
'            If RS1.State = 1 Then RS1.Close
'            RS1.Open "select * from opa", con, adOpenKeyset, adLockOptimistic
'            RS1.AddNew
'                RS1("act_idy") = Trim(RS("dbt_idy") & "")
'                RS1("opa_dbt") = Val(RS("txn_amt") & "")
'                RS1("opa_crt") = 0
'                RS1("fyr_dat") = TO_DAT
'            RS1.Update
'        End If
'        If RS1.State = 1 Then RS1.Close
'        RS1.Open "SELECT * FROM opa where act_idy like '" & Trim(RS("CRt_idy")) & "'", con, adOpenKeyset, adLockOptimistic
'        If RS1.RecordCount > 0 Then
'            RS1("opa_crt") = Val(RS1("opa_crt")) + Val(RS("txn_amt") & "")
'            RS1.Update
'        Else
'            If RS1.State = 1 Then RS1.Close
'            RS1.Open "select * from opa", con, adOpenKeyset, adLockOptimistic
'            RS1.AddNew
'                RS1("act_idy") = Trim(RS("crt_idy") & "")
'                RS1("opa_crt") = Val(RS("txn_amt") & "")
'                RS1("opa_dbt") = 0
'                RS1("fyr_dat") = TO_DAT
'            RS1.Update
'        End If
'        RS.MoveNext
'    Next
'    End If
'    If RS.State = 1 Then RS.Close
'    If RS1.State = 1 Then RS1.Close
'    RS.Open "select * from inl where inv_dat <=ctod('" & TO_DAT & "') and txn_typ in('LSL','SAL','PRT','RSI','STO','PCS')", con, adOpenKeyset, adLockOptimistic
'    If RS.RecordCount > 0 Then
'        For A = 1 To RS.RecordCount
'            If RS1.State = 1 Then RS1.Close
'            RS1.Open "SELECT * FROM opa where act_idy like '" & Trim(RS("CSt_idy")) & "'", con, adOpenKeyset, adLockOptimistic
'            If RS1.RecordCount > 0 Then
'                RS1("opa_dbt") = Val(RS1("opa_dbt")) + Val(RS("net_amt") & "")
'                RS1.Update
'            Else
'                If RS1.State = 1 Then RS1.Close
'                RS1.Open "select * from opa", con, adOpenKeyset, adLockOptimistic
'                RS1.AddNew
'                    RS1("act_idy") = Trim(RS("cst_idy") & "")
'                    RS1("opa_dbt") = Val(RS("net_amt") & "")
'                    RS1("opa_crt") = 0
'                    RS1("fyr_dat") = TO_DAT
'                RS1.Update
'            End If
'            If check("opa", "act_idy", Trim(RS("act_idy") & "")) = True Then
'                con.Execute "Update opa where act_idy like '" & Trim(RS("act_idy")) & "' set opa_crt = opa_crt+ " & Val(RS("net_amt") & "")
'            Else
'                If RS1.State = 1 Then RS1.Close
'                RS1.Open "select * from opa", con, adOpenKeyset, adLockOptimistic
'                RS1.AddNew
'                    RS1("act_idy") = Trim(RS("act_idy") & "")
'                    RS1("opa_dbt") = 0
'                    RS1("opa_crt") = Val(RS("net_amt") & "")
'                    RS1("fyr_dat") = TO_DAT
'                RS1.Update
'            End If
'            RS.MoveNext
'        Next
'    End If
'    If RS.State = 1 Then RS.Close
'    If RS1.State = 1 Then RS1.Close
'    RS.Open "select * from inl where inv_dat <=ctod('" & TO_DAT & "') and txn_typ in('PUR','SRN','RSR','STI','OPS')", con, adOpenKeyset, adLockOptimistic
'    If RS.RecordCount > 0 Then
'        For A = 1 To RS.RecordCount
'            If check("opa", "act_idy", Trim(RS("cst_idy") & "")) = True Then
'                con.Execute "Update opa where act_idy like '" & Trim(RS("cst_idy")) & "' set opa_CRt=opa_CRT+" & Val(RS("net_amt") & "")
'            Else
'                If RS1.State = 1 Then RS1.Close
'                RS1.Open "select * from opa", con, adOpenKeyset, adLockOptimistic
'                RS1.AddNew
'                    RS1("act_idy") = Trim(RS("cst_idy") & "")
'                    RS1("opa_CRt") = Val(RS("net_amt") & "")
'                    RS1("opa_DBT") = 0
'                    RS1("fyr_dat") = TO_DAT
'                RS1.Update
'            End If
'            If check("opa", "act_idy", Trim(RS("act_idy") & "")) = True Then
'                con.Execute "Update opa where act_idy like '" & Trim(RS("act_idy")) & "' set opa_DBt=opa_DBT+" & Val(RS("net_amt") & "")
'            Else
'                If RS1.State = 1 Then RS1.Close
'                RS1.Open "select * from opa", con, adOpenKeyset, adLockOptimistic
'                RS1.AddNew
'                    RS1("act_idy") = Trim(RS("act_idy") & "")
'                    RS1("opa_CRt") = 0
'                    RS1("opa_DBt") = Val(RS("net_amt") & "")
'                    RS1("fyr_dat") = TO_DAT
'                RS1.Update
'            End If
'            RS.MoveNext
'        Next
'    End If
'End If
End Sub

Public Sub PAL()
Dim MLOS As Long, MPRO As Long
Dim RS As New ADODB.Recordset
If RS.State = 1 Then RS.Close
Dim RS1 As New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
Dim rs2 As New ADODB.Recordset
If rs2.State = 1 Then rs2.Close
RS.Open "select * from opa", CON, adOpenKeyset, adLockOptimistic
For A = 1 To RS.RecordCount
    If RS1.State = 1 Then RS1.Close
    RS1.Open "select * from act WHERE act_idy like '" & Trim(RS("act_idy") & "") & "'", CON, adOpenKeyset, adLockOptimistic
    If RS1.RecordCount > 0 Then
        If rs2.State = 1 Then rs2.Close
        rs2.Open "select * from typ where typ_def like '" & Trim(RS1("grp_clf") & "") & "'", CON, adOpenKeyset, adLockOptimistic
        If rs2.RecordCount > 0 Then
            If Trim(rs2("typ_fac") & "") = "D" And Trim(rs2("typ_clf") & "") = "PL" Then
                MLOS = MLOS + Val(RS("opa_dbt"))
            ElseIf Trim(rs2("typ_fac") & "") = "C" And Trim(rs2("typ_clf") & "") = "PL" Then
                MPRO = MPRO + Val(RS("opa_crt"))
            End If
        End If
    End If
    RS.MoveNext
Next
CON.Execute "update dir set dir_als='" & MLOS & "', dir_msg='" & MPRO & "' where dir_idy like 'PRO' and dir_idx like 'LOS'"
End Sub

Public Function P_SUB_UNT(QTY As String, SUBQTY) As String
A = InStr(Trim(QTY), ".")
If A > 0 Then
    EDECNUM = Mid(QTY, A + 1, 3)
    NUM1 = Mid(QTY, 1, A - 1)
    P_SUB_UNT = Val(Val(NUM1) * Val(SUBQTY)) + Val(EDECNUM)
End If
End Function

Public Function P_QTY_UNT(M_QTY As String, S_QTY) As String
'r = Round((Val(QTY) Mod Val(SUBQTY)), 3)       'remainde
'v = Val(QTY) \ Val(SUBQTY)   'quotient
'If Len(Trim(r)) = 1 Then
'    P_QTY_UNT = v & ".00" & r
'ElseIf Len(Trim(r)) = 2 Then
'    P_QTY_UNT = v & ".0" & r
'ElseIf Len(Trim(r)) = 3 Then
'    P_QTY_UNT = v & "." & r
'End If
'''''
If Val(M_QTY) < 0 Then
    M_CHK = True
Else
    M_CHK = False
End If
M_QTY = Abs(M_QTY)
If S_QTY <> 1 Then
    M_MOD = Val(M_QTY) Mod S_QTY
    A = (M_QTY - M_MOD)
    b = Val(A) / Val(S_QTY)
    c = Val(b) + (M_MOD / 1000)
Else
    c = M_QTY
End If
If M_CHK = True Then c = c * -1
P_QTY_UNT = QTY(Trim(c))
End Function

Public Sub M_HELP()
M_PATH = App.Path & "\CONTHELP.CHM"
If company.SysInfo1.OSPlatform = 1 Then
    Shell "C:\WINDOWS\HH.EXE " & M_PATH, vbMaximizedFocus
ElseIf company.SysInfo1.OSPlatform = 2 Then
    Shell "C:\WINNT\HH.EXE " & M_PATH, vbMaximizedFocus
Else
    Shell "C:\WINDOWS\HH.EXE " & M_PATH, vbMaximizedFocus
End If
End Sub

Public Function Openreport(ReportLibrary As String, reportname As String, query As Integer, filter As String, destination As Integer)
With M_MDI.RRReport1
    .ReportLibrary = Trim(ReportLibrary)
    .reportname = Trim(reportname)
    .Printer = "?"
    .query = Val(query)
    .filter = filter
    .destination = Val(destination)
    .RunReport 1
End With
End Function

Public Sub Openscopereport(RP1 As String, REPNAME As String, qry As Integer, SCP As Integer, LOW As String, HIGH As String, filter As String, destination As Integer)
With M_MDI.RRReport1
    .ReportLibrary = RP1
    .reportname = REPNAME
    .query = qry
    .scope = SCP
    .LowScope = LOW
    .HighScope = HIGH
    .filter = filter
    .destination = destination
    .Printer = "?"
    .RunReport 1
End With
End Sub

Public Function OpenDocument(ReportLibrary As String, reportname As String, scope As Integer, loscope As String, hiscope As String, destination As Integer, mindex As String)
On Error GoTo ERR_TRAP
Dim RSdir As New ADODB.Recordset
With M_MDI.RRReport1
    .ReportLibrary = Trim(ReportLibrary)
    If reportname = "Invoice" Then
        RSdir.Open "select * from dir where dir_idy like 'AUT' and dir_idx like 'IND'", CON, adOpenKeyset, adLockOptimistic
        If RSdir.RecordCount > 0 Then
            If Trim(RSdir("DIR_MSG")) = "" Then
                .reportname = "Invoice"
            Else
                .reportname = Trim(RSdir("DIR_MSG"))
            End If
        End If
    Else
        .reportname = Trim(reportname)
    End If
    .Printer = "?"
    .scope = Val(scope)
    .MasterIndex = Trim(mindex)
    .LowScope = loscope
    .HighScope = hiscope
    .destination = Val(destination)
    .RunReport 1
End With
Exit Function
ERR_TRAP:
MsgBox err.Description
End Function

Public Sub UPD_STK()
M_FRM_UPD.Show 1
'If MsgBox("Update Stocks..", vbQuestion + vbYesNo) = vbYes Then
'    Dim REC_SET_I As New ADODB.Recordset
'    Dim REC_SET_P As New ADODB.Recordset
'    con.Execute "UPDATE PRD SET CUR_LEV = 0"
'    REC_SET_I.Open "SELECT * FROM INV,PRD WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND PRD.SUB_UNT <> 1", con, adOpenKeyset, adLockReadOnly
'    For a = 1 To REC_SET_I.RecordCount
'            If REC_SET_I("TXN_TYP") = "OPS" Or REC_SET_I("TXN_TYP") = "PUR" Or REC_SET_I("TXN_TYP") = "SRN" Or REC_SET_I("TXN_TYP") = "STO" Or REC_SET_I("TXN_TYP") = "RSR" Or REC_SET_I("TXN_TYP") = "PCR" Then
'                con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET_I("PRD_IDY")) & "' SET CUR_LEV = " & ADDQTY(REC_SET_I("PRD_QTY"), REC_SET_I("PRD_IDY"))
'            Else
'                con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET_I("PRD_IDY")) & "' SET CUR_LEV = " & CONQTY(REC_SET_I("PRD_QTY"), REC_SET_I("PRD_IDY"))
'            End If
'        REC_SET_I.MoveNext
'    Next
'
'    If REC_SET_I.State = 1 Then REC_SET_I.Close
'    REC_SET_I.Open "SELECT INV.PRD_IDY,SUM(INV.PRD_QTY) FROM INV,PRD WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND PRD.SUB_UNT = 1 AND TXN_TYP IN('OPS','PUR','SRN','STO','RSR') GROUP BY INV.PRD_IDY", con, adOpenKeyset, adLockReadOnly
'    For a = 1 To REC_SET_I.RecordCount
'        con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET_I("PRD_IDY")) & "' SET CUR_LEV = CUR_LEV + " & Val(REC_SET_I(1))
'        REC_SET_I.MoveNext
'    Next
'
'    If REC_SET_I.State = 1 Then REC_SET_I.Close
'    REC_SET_I.Open "SELECT INV.PRD_IDY,SUM(PRD_QTY) FROM INV,PRD WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND PRD.SUB_UNT = 1 AND TXN_TYP IN('LSL','SAL','PRT','STI','RSI') GROUP BY INV.PRD_IDY ", con, adOpenKeyset, adLockReadOnly
'    For a = 1 To REC_SET_I.RecordCount
'        con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET_I("PRD_IDY")) & "' SET CUR_LEV = CUR_LEV - " & Val(REC_SET_I(1))
'        REC_SET_I.MoveNext
'    Next
'End If
End Sub

Public Sub N_UPD_STK()
If MsgBox("Update Stocks..", vbQuestion + vbYesNo) = vbYes Then
    Dim REC_SET As New ADODB.Recordset
    Dim REC_SET1 As New ADODB.Recordset
    
    CON.Execute "UPDATE PRD SET CUR_LEV = 0"
    REC_SET.Open "SELECT * FROM INV where txn_typ in('OPS','PUR','SRN','STO','RSR','PCR')", CON, adOpenKeyset, adLockOptimistic
    If REC_SET.RecordCount > 0 Then
       For A = 1 To REC_SET.RecordCount
                If REC_SET1.State = 1 Then REC_SET1.Close
                REC_SET1.Open "SELECT * FROM PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET("PRD_IDY")) & "'", CON, adOpenKeyset, adLockOptimistic
                If REC_SET1.RecordCount > 0 Then
                    PRD_SUB = P_SUB_UNT(QTY(REC_SET1("CUR_LEV") & ""), REC_SET1("SUB_UNT"))
                    INV_SUB = P_SUB_UNT(QTY(REC_SET("PRD_QTY") & ""), REC_SET1("SUB_UNT"))
                    REC_SET1("CUR_LEV") = P_QTY_UNT(Val(PRD_SUB) + Val(INV_SUB), REC_SET1("SUB_UNT"))
                    REC_SET1.Update
                End If
                REC_SET.MoveNext
        Next
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM INV where txn_typ in('LSL','SAL','PRT','STI','RSI','PCS')", CON, adOpenKeyset, adLockOptimistic
        For A = 1 To REC_SET.RecordCount
            If REC_SET1.State = 1 Then REC_SET1.Close
            REC_SET1.Open "SELECT * FROM PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET("PRD_IDY")) & "'", CON, adOpenKeyset, adLockOptimistic
            If REC_SET1.RecordCount > 0 Then
                PRD_SUB = P_SUB_UNT(QTY(REC_SET1("CUR_LEV") & ""), REC_SET1("SUB_UNT"))
                INV_SUB = P_SUB_UNT(QTY(REC_SET("PRD_QTY") & ""), REC_SET1("SUB_UNT"))
                REC_SET1("CUR_LEV") = P_QTY_UNT(Val(PRD_SUB) - Val(INV_SUB), REC_SET1("SUB_UNT"))
                REC_SET1.Update
            End If
            REC_SET.MoveNext
        Next
    End If
End If
End Sub

Public Sub STL_LED(FRM_DAT As Date, to_dat As Date, qry As String)
M_FRM_PPQ.MousePointer = vbHourglass
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
CON.Execute "delete from stl"
If RS.State = 1 Then RS.Close
If RS1.State = 1 Then RS1.Close
RS.Open "select * from stl", CON, adOpenKeyset, adLockOptimistic
RS1.Open "select * from inl,INV,PRD WHERE STK_UPD ='Y' AND PRD.PRD_IDY LIKE INV.PRD_IDY AND inv.inv_idy like inl.inv_idy and INV.TXN_TYP IN ('PUR','SRN','STI','DMR','OPS')  and INV.inv_dat >= ctod('" & FRM_DAT & "') AND INV.INV_DAT <=CTOD('" & to_dat & "')" & Trim(qry), CON, adOpenKeyset, adLockOptimistic
For A = 1 To RS1.RecordCount
    RS.AddNew
    RS("inv_idy") = Trim(RS1("inv_idy"))
    RS("inv_dat") = Trim(RS1("inv_dat"))
    RS("prd_idy") = Trim(RS1("prd_idy"))
    RS("ISS_qty") = 0
    RS("RCT_qty") = Trim(RS1("prd_qty"))
    RS("txn_typ") = Trim(RS1("txn_typ"))
    RS("cst_idy") = Trim(RS1("cst_idy"))
    RS("TXN_RMK") = Trim(RS1("TXN_RMK") & "")
    RS.Update
    RS1.MoveNext
Next
If RS.State = 1 Then RS.Close
If RS1.State = 1 Then RS1.Close
RS.Open "select * from stl", CON, adOpenKeyset, adLockOptimistic
RS1.Open "select inv.* from INV,PRD,PGR WHERE STK_UPD ='Y' AND PRD.PRD_IDY LIKE INV.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND inv.TXN_TYP IN ('SAL','PRT','STO') AND inv.inv_dat >= ctod('" & FRM_DAT & "') AND INV.INV_DAT <=CTOD('" & to_dat & "')" & Trim(qry), CON, adOpenKeyset, adLockOptimistic
For A = 1 To RS1.RecordCount
    RS.AddNew
    RS("inv_idy") = Trim(RS1("inv_idy"))
    RS("inv_dat") = Trim(RS1("inv_dat"))
    RS("prd_idy") = Trim(RS1("prd_idy"))
    If Trim(RS1("stk_typ")) = "Salable" Or Trim(RS1("stk_typ")) = "Free" Then
        RS("rct_qty") = 0
        RS("iss_qty") = Trim(RS1("prd_qty"))
    ElseIf Trim(RS1("stk_typ")) = "TakeBack" Or Trim(RS1("stk_typ")) = "Damage" Then
        RS("iss_qty") = 0
        RS("rct_qty") = Trim(RS1("prd_qty"))
    End If
    RS("txn_typ") = Trim(RS1("txn_typ"))
    If RS("txn_typ") = "FRS" Or RS("TXN_TYP") = "FRC" Then
        If rs2.State = 1 Then rs2.Close
        rs2.Open "SELECT CST_IDY,TXN_RMK FROM INL WHERE INV_IDY LIKE '" & Trim(RS1("INV_IDY")) & "'", CON, adOpenKeyset, adLockOptimistic
    Else
        If rs2.State = 1 Then rs2.Close
        rs2.Open "SELECT CST_IDY,TXN_RMK FROM INL WHERE INV_IDY LIKE '" & Trim(RS1("INV_IDY")) & "' AND TXN_TYP LIKE '" & Trim(RS1("TXN_TYP")) & "'", CON, adOpenKeyset, adLockOptimistic
    End If
    If rs2.RecordCount > 0 Then
        RS("cst_idy") = Trim(rs2("cst_idy"))
        RS("TXN_RMK") = Trim(rs2("TXN_RMK") & " ")
    End If
    RS.Update
    RS1.MoveNext
Next
'OPS UPDATION
CON.Execute "delete from OPS"
If RS.State = 1 Then RS.Close
If RS1.State = 1 Then RS1.Close
RS.Open "select * from OPS", CON, adOpenKeyset, adLockOptimistic
RS1.Open "select * from INV,PRD,PGR WHERE STK_UPD ='Y' AND PRD.PRD_IDY LIKE INV.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND TXN_TYP IN ('PUR','SRN','STI','DMR','OPS')  and inv_dat < ctod('" & FRM_DAT & "')" & Trim(qry), CON, adOpenKeyset, adLockOptimistic
For A = 1 To RS1.RecordCount
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM OPS WHERE PRD_IDY LIKE '" & Trim(RS1("PRD_IDY")) & "'", CON, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
        R_QTY = Val(MTS(Val(RS("PRD_QTY") & ""), RS1("SUB_UNT"))) + Val(MTS(RS1("PRD_QTY"), RS1("SUB_UNT")))
        RS("PRD_QTY") = STM(Val(R_QTY), RS1("SUB_UNT"))
        RS.Update
    Else
        RS.AddNew
        RS("prd_idy") = Trim(RS1("prd_idy"))
        RS("PRD_qty") = Trim(RS1("prd_qty"))
        RS.Update
    End If
    RS1.MoveNext
Next
If RS.State = 1 Then RS.Close
If RS1.State = 1 Then RS1.Close
RS1.Open "select * from INV,PRD,PGR WHERE STK_UPD ='Y' AND PRD.PRD_IDY LIKE INV.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND inv.TXN_TYP IN ('SAL','PRT','STO') AND inv.inv_dat < ctod('" & FRM_DAT & "')" & Trim(qry), CON, adOpenKeyset, adLockOptimistic
For A = 1 To RS1.RecordCount
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM OPS WHERE PRD_IDY LIKE '" & Trim(RS1("PRD_IDY")) & "'", CON, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
        R_QTY = Val(MTS(Val(RS("PRD_QTY") & ""), RS1("SUB_UNT"))) - Val(MTS(RS1("PRD_QTY"), RS1("SUB_UNT")))
        RS("PRD_QTY") = STM(Val(R_QTY), RS1("SUB_UNT"))
        RS.Update
    Else
        RS.AddNew
        RS("prd_idy") = Trim(RS1("prd_idy"))
        RS("PRD_qty") = Trim(RS1("prd_qty"))
        RS.Update
    End If
    RS1.MoveNext
Next
M_FRM_PPQ.MousePointer = vbDefault
Openreport CURDIR & "\Report.rp1", "Stock Ledger", 0, Trim(SAL_QRY), 6
End Sub
Public Sub FAupdateCLG(FRM_DAT1 As Date, TO_DATE1 As Date, grpid As String)
FRM_DAT = FRM_DAT1
TO_DATE = TO_DATE1
M_FRM_APD.Tag = "FA_CLG"
M_FRM_APD.Show 1
End Sub
Public Sub FAupdate(FRM_DAT1 As Date, TO_DATE1 As Date, grpid As String)
'con.Execute "delete from led"
'Dim rs As New ADODB.Recordset
'Dim RS1 As New ADODB.Recordset
'If rs.State = 1 Then rs.Close
'If RS1.State = 1 Then RS1.Close
'TRAIL_BALANCE FRM_DAT
''RS.Open "select * from vch WHERE TXN_TYP NOT LIKE 'OPN' and txn_dat >= ctod('" & FRM_DAT & "') AND TXN_DAT <= CTOD('" & to_date & "')", con, adOpenKeyset, adLockOptimistic
'rs.Open "select * from vch WHERE txn_dat >= ctod('" & FRM_DAT & "') AND TXN_DAT <= CTOD('" & to_date & "')", con, adOpenKeyset, adLockOptimistic
'RS1.Open "select * from led", con, adOpenKeyset, adLockOptimistic
'For A = 1 To rs.RecordCount
'    RS1.AddNew
'        RS1(0) = Trim(rs("txn_idy") & "")
'        RS1(1) = rs("txn_dat")
'        RS1(2) = rs("crt_idy")
'        RS1(4) = rs("txn_amt")
'        RS1(3) = 0
'        RS1(5) = rs("txn_typ")
'        RS1(6) = Trim(rs("txn_rmk") & "")
'        RS1(7) = Trim(rs("dbt_idy") & "")
'    RS1.Update
'    RS1.AddNew
'        RS1(0) = Trim(rs("txn_idy") & "")
'        RS1(1) = rs("txn_dat")
'        RS1(2) = rs("dbt_idy")
'        RS1(4) = 0
'        RS1(3) = Val(Trim(rs("txn_amt") & ""))
'        RS1(5) = rs("txn_typ")
'        RS1(6) = Trim(rs("txn_rmk") & "")
'        RS1(7) = Trim(rs("crt_idy") & "")
'    RS1.Update
'    rs.MoveNext
'Next
'If rs.State = 1 Then rs.Close
'If RS1.State = 1 Then RS1.Close
'rs.Open "select * from inl where inv_dat >=ctod('" & FRM_DAT & "') AND inv_dat <=ctod('" & to_date & "') AND TXN_TYP IN ('SAL','PRT','LSL','STI','RSI','STO','PCS')", con, adOpenKeyset, adLockOptimistic
'RS1.Open "select * from led", con, adOpenKeyset, adLockOptimistic
'For A = 1 To rs.RecordCount
'    RS1.AddNew
'    RS1(0) = Trim(rs("inv_idy") & "")
'    RS1(1) = rs("inv_dat")
'    RS1(2) = Trim(rs("cst_idy"))
'    RS1(3) = rs("net_amt")
'    RS1(4) = 0
'
'    RS1(5) = Trim(rs("txn_typ"))
'    RS1(6) = Trim(rs("txn_rmk") & "")
'    RS1(7) = Trim(rs("act_idy"))
'    RS1.Update
'
'    RS1.AddNew
'    RS1(0) = Trim(rs("inv_idy") & "")
'    RS1(1) = rs("inv_dat")
'    RS1(2) = Trim(rs("act_idy"))
'    RS1(3) = 0
'    RS1(4) = Val(Trim(rs("net_amt") & ""))
'    RS1(5) = Trim(rs("txn_typ"))
'    RS1(6) = Trim(rs("txn_rmk") & "")
'    RS1(7) = Trim(rs("cst_idy"))
'    RS1.Update
'    rs.MoveNext
'Next
'If rs.State = 1 Then rs.Close
'If RS1.State = 1 Then RS1.Close
'rs.Open "select * from inl where inv_dat >=ctod('" & FRM_DAT & "') AND inv_dat <=ctod('" & to_date & "') AND TXN_TYP IN ('PUR','SRN','ST0','RSR','OPS','STI','PCR')", con, adOpenKeyset, adLockOptimistic
'RS1.Open "select * from led", con, adOpenKeyset, adLockOptimistic
'For A = 1 To rs.RecordCount
'    RS1.AddNew
'    RS1(0) = Trim(rs("inv_idy") & "")
'    RS1(1) = rs("inv_dat")
'    RS1(2) = Trim(rs("cst_idy"))
'    RS1(4) = rs("net_amt")
'    RS1(3) = 0
'    RS1(5) = Trim(rs("txn_typ"))
'    RS1(6) = Trim(rs("txn_rmk") & "")
'    RS1(7) = Trim(rs("act_idy"))
'    RS1.Update
'
'    RS1.AddNew
'    RS1(0) = Trim(rs("inv_idy") & "")
'    RS1(1) = rs("inv_dat")
'    RS1(2) = Trim(rs("act_idy"))
'    RS1(4) = 0
'    RS1(3) = Val(Trim(rs("net_amt") & ""))
'    RS1(5) = Trim(rs("txn_typ"))
'    RS1(6) = Trim(rs("txn_rmk") & "")
'    RS1(7) = Trim(rs("cst_idy"))
'    RS1.Update
'    rs.MoveNext
'Next
FRM_DAT = FRM_DAT1
TO_DATE = TO_DATE1
M_FRM_APD.Tag = "FA"
M_FRM_APD.Show 1
End Sub

Public Sub O_UPD_STK()
If MsgBox("Update Stocks..", vbExclamation + vbYesNo) = vbYes Then
CON.Execute "UPDATE PRD SET CUR_LEV = 0"
Dim RSI As New ADODB.Recordset
If RSI.State = 1 Then RSI.Close
RSI.Open "SELECT INV.TXN_TYP,INV.PRD_QTY,INV.PRD_IDY,PRD.SUB_UNT FROM INV,PRD WHERE INV.TXN_TYP IN ('OPS','PUR','SRN','STI','RSR') AND PRD.PRD_IDY LIKE INV.PRD_IDY", CON, adOpenKeyset, adLockOptimistic
Do While Not RSI.EOF
    If Val(RSI(3)) <> 1 Then
        CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= " & ADDQTY(RSI(1), RSI(2))
    Else
        CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= CUR_LEV +" & RSI(1)
    End If
    RSI.MoveNext
Loop

If RSI.State = 1 Then RSI.Close
RSI.Open "SELECT INV.TXN_TYP,INV.PRD_QTY,INV.PRD_IDY,PRD.SUB_UNT FROM INV,PRD WHERE INV.TXN_TYP LIKE 'LSL' AND PRD.PRD_IDY LIKE INV.PRD_IDY", CON, adOpenKeyset, adLockOptimistic
Do While Not RSI.EOF
    If Val(RSI(3)) <> 1 Then
        CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= " & CONQTY(RSI(1), RSI(2))
    Else
        CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= CUR_LEV- " & RSI(1)
    End If
    RSI.MoveNext
Loop

If RSI.State = 1 Then RSI.Close
RSI.Open "SELECT INV.TXN_TYP,INV.PRD_QTY,INV.PRD_IDY,PRD.SUB_UNT FROM INV,PRD WHERE INV.TXN_TYP IN ('SAL','RSI','STO','PRT') AND PRD.PRD_IDY LIKE INV.PRD_IDY", CON, adOpenKeyset, adLockOptimistic
Do While Not RSI.EOF
    If Val(RSI(3)) <> 1 Then
        CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= " & CONQTY(RSI(1), RSI(2))
    Else
        CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= CUR_LEV -" & RSI(1)
    End If
    RSI.MoveNext
Loop

End If
End Sub

Public Function MTS(M_QTY As Double, S_QTY As Integer) As Long
If Val(S_QTY) <> 1 Then
    A = Int(M_QTY) * Val(S_QTY)
    b = M_QTY - Int(M_QTY)
    b = b * 1000
    c = Val(A) + Val(b)
Else
    c = M_QTY
End If
MTS = c
End Function

Public Function STM(M_QTY As Double, S_QTY As Integer) As Double
If M_QTY < 0 Then
    M_CHK = True
Else
    M_CHK = False
End If
M_QTY = Abs(M_QTY)
If S_QTY <> 1 Then
    M_MOD = Val(M_QTY) Mod S_QTY
    A = (M_QTY - M_MOD)
    b = Val(A) / Val(S_QTY)
    c = Val(b) + (M_MOD / 1000)
Else
    c = M_QTY
End If
If M_CHK = True Then c = c * -1
STM = c
End Function

Public Sub createcompany(Destdir As String)
If Dir(App.Path & "\bkp", vbDirectory) = "" Then
MkDir App.Path & "\bkp"
Destdir = App.Path & "\bkp"
CopyFile App.Path & "\act.dbf", Destdir & "\act.dbf", 0
CopyFile App.Path & "\att.dbf", Destdir & "\att.dbf", 0
CopyFile App.Path & "\asm.dbf", Destdir & "\asm.dbf", 0
CopyFile App.Path & "\wek.dbf", Destdir & "\wek.dbf", 0
CopyFile App.Path & "\tgt.dbf", Destdir & "\tgt.dbf", 0
CopyFile App.Path & "\tgS.dbf", Destdir & "\tgS.dbf", 0
CopyFile App.Path & "\tgB.dbf", Destdir & "\tgB.dbf", 0
CopyFile App.Path & "\tgC.dbf", Destdir & "\tgC.dbf", 0
CopyFile App.Path & "\ast.dbf", Destdir & "\ast.dbf", 0
CopyFile App.Path & "\ind.exe", Destdir & "\ind.exe", 0
CopyFile App.Path & "\lds.dbf", Destdir & "\lds.dbf", 0
CopyFile App.Path & "\sch.dbf", Destdir & "\sch.dbf", 0
CopyFile App.Path & "\pnl.dbf", Destdir & "\pnl.dbf", 0
CopyFile App.Path & "\pnv.dbf", Destdir & "\pnv.dbf", 0
CopyFile App.Path & "\log.dbf", Destdir & "\log.dbf", 0
CopyFile App.Path & "\rri.dbf", Destdir & "\rri.dbf", 0
CopyFile App.Path & "\lev.dbf", Destdir & "\lev.dbf", 0
CopyFile App.Path & "\lnv.dbf", Destdir & "\lnv.dbf", 0
CopyFile App.Path & "\brd.dbf", Destdir & "\brd.dbf", 0
CopyFile App.Path & "\van.dbf", Destdir & "\van.dbf", 0
CopyFile App.Path & "\ord.dbf", Destdir & "\ord.dbf", 0
CopyFile App.Path & "\irt.dbf", Destdir & "\irt.dbf", 0
CopyFile App.Path & "\agd.dbf", Destdir & "\agd.dbf", 0
CopyFile App.Path & "\loc.dbf", Destdir & "\loc.dbf", 0
CopyFile App.Path & "\cat.dbf", Destdir & "\cat.dbf", 0
CopyFile App.Path & "\cld.dbf", Destdir & "\cld.dbf", 0
CopyFile App.Path & "\dir.dbf", Destdir & "\dir.dbf", 0
CopyFile App.Path & "\dsu.dbf", Destdir & "\dsu.dbf", 0
CopyFile App.Path & "\fun.dbf", Destdir & "\fun.dbf", 0
CopyFile App.Path & "\grp.dbf", Destdir & "\grp.dbf", 0
CopyFile App.Path & "\idy.dbf", Destdir & "\idy.dbf", 0
CopyFile App.Path & "\ine.dbf", Destdir & "\ine.dbf", 0
CopyFile App.Path & "\inv.dbf", Destdir & "\inv.dbf", 0
CopyFile App.Path & "\inl.dbf", Destdir & "\inl.dbf", 0
CopyFile App.Path & "\led.dbf", Destdir & "\led.dbf", 0
CopyFile App.Path & "\mar.dbf", Destdir & "\mar.dbf", 0
CopyFile App.Path & "\mbr.dbf", Destdir & "\mbr.dbf", 0
CopyFile App.Path & "\mcl.dbf", Destdir & "\mcl.dbf", 0
CopyFile App.Path & "\npo.dbf", Destdir & "\npo.dbf", 0
CopyFile App.Path & "\opa.dbf", Destdir & "\opa.dbf", 0
CopyFile App.Path & "\ops.dbf", Destdir & "\ops.dbf", 0
CopyFile App.Path & "\pgr.dbf", Destdir & "\pgr.dbf", 0
CopyFile App.Path & "\ppc.dbf", Destdir & "\ppc.dbf", 0
CopyFile App.Path & "\prd.dbf", Destdir & "\prd.dbf", 0
CopyFile App.Path & "\pre.dbf", Destdir & "\pre.dbf", 0
CopyFile App.Path & "\csr.dbf", Destdir & "\csr.dbf", 0
CopyFile App.Path & "\prt.dbf", Destdir & "\prt.dbf", 0
CopyFile App.Path & "\smn.dbf", Destdir & "\smn.dbf", 0
CopyFile App.Path & "\sta.dbf", Destdir & "\sta.dbf", 0
CopyFile App.Path & "\rtp.dbf", Destdir & "\rtp.dbf", 0
CopyFile App.Path & "\stl.dbf", Destdir & "\stl.dbf", 0
CopyFile App.Path & "\sts.dbf", Destdir & "\sts.dbf", 0
CopyFile App.Path & "\top.dbf", Destdir & "\top.dbf", 0
CopyFile App.Path & "\tpd.dbf", Destdir & "\tpd.dbf", 0
CopyFile App.Path & "\typ.dbf", Destdir & "\typ.dbf", 0
CopyFile App.Path & "\usr.dbf", Destdir & "\usr.dbf", 0
CopyFile App.Path & "\vch.dbf", Destdir & "\vch.dbf", 0
CopyFile App.Path & "\yer.dbf", Destdir & "\yer.dbf", 0
CopyFile App.Path & "\eff.dbc", Destdir & "\eff.dbc", 0
CopyFile App.Path & "\eff.dct", Destdir & "\eff.dct", 0
CopyFile App.Path & "\eff.dcx", Destdir & "\eff.dcx", 0
CopyFile App.Path & "\mar.cdx", Destdir & "\mar.cdx", 0
CopyFile App.Path & "\brd.cdx", Destdir & "\brd.cdx", 0
CopyFile App.Path & "\irt.cdx", Destdir & "\irt.cdx", 0
CopyFile App.Path & "\lds.cdx", Destdir & "\lds.cdx", 0
CopyFile App.Path & "\lev.cdx", Destdir & "\lev.cdx", 0
CopyFile App.Path & "\lnv.cdx", Destdir & "\lnv.cdx", 0
CopyFile App.Path & "\mcl.cdx", Destdir & "\mcl.cdx", 0
CopyFile App.Path & "\ord.cdx", Destdir & "\ord.cdx", 0
CopyFile App.Path & "\pnl.cdx", Destdir & "\pnl.cdx", 0
CopyFile App.Path & "\pnv.cdx", Destdir & "\pnv.cdx", 0
CopyFile App.Path & "\sch.cdx", Destdir & "\sch.cdx", 0
CopyFile App.Path & "\van.cdx", Destdir & "\van.cdx", 0
CopyFile App.Path & "\act.cdx", Destdir & "\act.cdx", 0
CopyFile App.Path & "\grp.cdx", Destdir & "\grp.cdx", 0
CopyFile App.Path & "\ops.cdx", Destdir & "\ops.cdx", 0
CopyFile App.Path & "\prd.cdx", Destdir & "\prd.cdx", 0
CopyFile App.Path & "\pgr.cdx", Destdir & "\pgr.cdx", 0
CopyFile App.Path & "\smn.cdx", Destdir & "\smn.cdx", 0
CopyFile App.Path & "\inv.cdx", Destdir & "\inv.cdx", 0
CopyFile App.Path & "\inl.cdx", Destdir & "\inl.cdx", 0
CopyFile App.Path & "\vch.cdx", Destdir & "\vch.cdx", 0
CopyFile App.Path & "\dir.cdx", Destdir & "\dir.cdx", 0
CopyFile App.Path & "\opa.cdx", Destdir & "\opa.cdx", 0
CopyFile App.Path & "\typ.cdx", Destdir & "\typ.cdx", 0
CopyFile App.Path & "\loc.cdx", Destdir & "\loc.cdx", 0
CopyFile App.Path & "\cat.cdx", Destdir & "\cat.cdx", 0
CopyFile App.Path & "\sta.cdx", Destdir & "\sta.cdx", 0
CopyFile App.Path & "\cld.cdx", Destdir & "\cld.cdx", 0
CopyFile App.Path & "\tpd.cdx", Destdir & "\tpd.cdx", 0
CopyFile App.Path & "\agd.cdx", Destdir & "\agd.cdx", 0
CopyFile App.Path & "\stl.cdx", Destdir & "\stl.cdx", 0
CopyFile App.Path & "\led.cdx", Destdir & "\led.cdx", 0
CopyFile App.Path & "\mbr.cdx", Destdir & "\mbr.cdx", 0
CopyFile App.Path & "\yer.cdx", Destdir & "\yer.cdx", 0

CopyFile App.Path & "\report.rp1", Destdir & "\report.rp1", 0
CopyFile App.Path & "\ledger.rp1", Destdir & "\ledger.rp1", 0
CopyFile App.Path & "\finance.rp1", Destdir & "\finance.rp1", 0
CopyFile App.Path & "\lists.rp1", Destdir & "\lists.rp1", 0
CopyFile App.Path & "\mir.rp1", Destdir & "\mir.rp1", 0
CopyFile App.Path & "\stock.rp1", Destdir & "\stock.rp1", 0
CopyFile App.Path & "\vouch.rp1", Destdir & "\vouch.rp1", 0
CopyFile App.Path & "\rep_lib.rp1", Destdir & "\rep_lib.rp1", 0
CopyFile App.Path & "\dos_vch.rp1", Destdir & "\dos_vch.rp1", 0
CopyFile App.Path & "\IND.EXE", Destdir & "\IND.EXE", 0
CopyFile App.Path & "\RRUNTIME.EXE", Destdir & "\RRUNTIME.EXE", 0
CopyFile App.Path & "\RRUN.EXE", Destdir & "\RRUN.EXE", 0
 CopyFile App.Path & "\Sales Register.rrw", Destdir & "\Sales Register.rrw", 0
CopyFile App.Path & "\RR.CNF", Destdir & "\RR.CNF", 0
End If
End Sub

Public Function CHKNUM(K As Integer) As Integer
If K > 47 And K < 58 Or K = 8 Or K = 46 Then
CHKNUM = K
Else
CHKNUM = 0
End If
End Function
Public Function CHGCASE(K As Integer) As Integer
If K > 96 And K < 123 Then
CHGCASE = K - 32
Else
CHGCASE = K
End If
End Function

Public Function PERCENT(AMT As String, PRC As String) As String
PERCENT = Round(Val(AMT) * (Val(PRC) / 100), 2)
End Function
Public Function PERCADD(AMT As String, PRC As String) As String
PCGADD = (Val(AMT) * (Val(PRC) / 100)) + AMT
End Function
Public Function PERCLESS(AMT As String, PRC As String) As String
PCGADD = AMT - (Val(AMT) * (Val(PRC) / 100))
End Function

Public Function GetValue(fldnames As String, tblname As String) As String
Dim RS As New ADODB.Recordset
RS.Open "select " & fldnames & " from " & tblname, CON, adOpenKeyset, adLockOptimistic
If RS.RecordCount > 0 Then
    GetValue = Trim(RS(0) & "")
End If
End Function
Public Sub TRAIL_BALANCECLG(TO_DAT1 As Date)
to_dat = TO_DAT1
Unload M_FRM_ASD
M_FRM_APD.Tag = "TB_CLG"
M_FRM_APD.Show 1
End Sub
Public Sub inltopnl()
On Error GoTo errh1
If RS.State = 1 Then RS.Close
If RS1.State = 1 Then RS1.Close
RS.Open "select * from inl", CN1, adOpenKeyset, adLockOptimistic
RS1.Open "select * from pnl", CN1, adOpenKeyset, adLockOptimistic
CN1.Execute "delete from pnl"
While Not RS.EOF
RS1.AddNew
For i = 0 To RS.Fields.Count - 1
fld = RS.Fields(i).name
RS1(fld) = RS(fld)
Next
RS1.Update
RS.MoveNext
RS1.CancelUpdate
Wend
'MsgBox "completed"
If RS.RecordCount > 0 Then
RS.MoveFirst
While Not RS.EOF
For i = 4 To RS.Fields.Count - 1
If RS.Fields(i).Type = adInteger Or RS.Fields(i).Type = adNumeric Or RS.Fields(i).Type = adBigInt Then
CN1.Execute "update pnl set " & RS.Fields(i).name & "=" & RS(i) & "where inv_idy like '" & RS("inv_idy") & "'"
ElseIf RS.Fields(i).Type = adDate Or RS.Fields(i).Type = adDBDate Or RS.Fields(i).Type = adDBTimeStamp Then
CN1.Execute "update pnl set " & RS.Fields(i).name & "=ctod('" & RS(i) & "') where inv_idy like '" & RS("inv_idy") & "'"
Else
CN1.Execute "update pnl set " & RS.Fields(i).name & "='" & RS(i) & "' where inv_idy like '" & RS("inv_idy") & "'"
End If
Next
RS.MoveNext
'MsgBox "comp"
Wend
End If
'MsgBox "Full"
If RS.State = 1 Then RS.Close
If RS1.State = 1 Then RS1.Close
Exit Sub
errh1:
If err.Number = 3265 Then
    Resume Next
ElseIf err.Number = -2147217887 Then
If RS1(fld).Type = adChar Or RS1(fld).Type = adVarChar Or RS1(fld).Type = adBSTR Or RS1(fld).Type = adLongVarChar Then
RS1(fld) = ""
Resume Next
ElseIf RS1(fld).Type = adNumeric Or RS1(fld).Type = adBigInt Or RS1(fld).Type = adInteger Or RS1(fld).Type = adSingle Or RS1(fld).Type = adDouble Then
RS1(fld) = 0
Resume Next
ElseIf RS1(fld).Type = adDate Or RS1(fld).Type = adDBDate Or RS1(fld).Type = adDBTimeStamp Then
RS1(fld) = Date
Resume Next
End If
ElseIf err.Number = -2147467259 Then
'fields does not exist in inl and not null
fnm = Trim(Mid(err.Description, 45, 7))
found = 0
For K = 1 To Len(fnm)
If Mid(Trim(fnm), K, 1) = "_" Then
found = 1
Exit For
End If
Next
If found = 1 Then
If RS1(fnm).Type = adChar Or RS1(fnm).Type = adVarChar Or RS1(fnm).Type = adBSTR Or RS1(fnm).Type = adLongVarChar Then
'cn1.Execute "update pnl set " & fnm & "='.'"
RS1(fnm) = ""
Resume Next
ElseIf RS1(fnm).Type = adNumeric Or RS1(fnm).Type = adBigInt Or RS1(fnm).Type = adInteger Or RS1(fnm).Type = adSingle Or RS1(fld).Type = adDouble Then
'cn1.Execute "update pnl set " & fnm & "= 0"
RS1(fnm) = 0
Resume Next
ElseIf RS1(fnm).Type = adDate Or RS1(fnm).Type = adDBDate Or RS1(fnm).Type = adDBTimeStamp Then
'cn1.Execute "update pnl set " & fnm & "=" & Date
RS1(fnm) = Date
Resume Next
End If
End If
ElseIf err.Number = -2147217900 Then
Resume Next
Else
MsgBox err.Description & " " & err.Number & "  " & fld
End If
End Sub
Public Sub invtopnv()
On Error GoTo errh2
If RS.State = 1 Then RS.Close
If RS1.State = 1 Then RS1.Close
RS.Open "select * from inv", CN1, adOpenKeyset, adLockOptimistic
RS1.Open "select * from pnv", CN1, adOpenKeyset, adLockOptimistic
CN1.Execute "delete from pnv"
While Not RS.EOF
RS1.AddNew
For i = 0 To RS.Fields.Count - 1
fld = RS.Fields(i).name
RS1(fld) = RS(fld)
Next
RS1.Update
RS.MoveNext
RS1.CancelUpdate
Wend
'MsgBox "completed"
If RS.RecordCount > 0 Then
RS.MoveFirst
While Not RS.EOF
For i = 3 To RS.Fields.Count - 1
If RS.Fields(i).Type = adInteger Or RS.Fields(i).Type = adNumeric Or RS.Fields(i).Type = adBigInt Then
CN1.Execute "update pnv set " & RS.Fields(i).name & "=" & RS(i) & "where inv_idy like '" & RS("inv_idy") & "'"
ElseIf RS.Fields(i).Type = adDate Or RS.Fields(i).Type = adDBDate Or RS.Fields(i).Type = adDBTimeStamp Then
CN1.Execute "update pnv set " & RS.Fields(i).name & "=ctod('" & RS(i) & "') where inv_idy like '" & RS("inv_idy") & "'"
Else
CN1.Execute "update pnv set " & RS.Fields(i).name & "='" & RS(i) & "' where inv_idy like '" & RS("inv_idy") & "'"
End If
Next
RS.MoveNext
'MsgBox "comp"
Wend
End If
'MsgBox "Full"
If RS.State = 1 Then RS.Close
If RS1.State = 1 Then RS1.Close
Exit Sub
errh2:
If err.Number = 3265 Then
    Resume Next
ElseIf err.Number = -2147217887 Then
If RS1(fld).Type = adChar Or RS1(fld).Type = adVarChar Or RS1(fld).Type = adBSTR Or RS1(fld).Type = adLongVarChar Then
RS1(fld) = ""
Resume Next
ElseIf RS1(fld).Type = adNumeric Or RS1(fld).Type = adBigInt Or RS1(fld).Type = adInteger Or RS1(fld).Type = adSingle Or RS1(fld).Type = adDouble Then
RS1(fld) = 0
Resume Next
ElseIf RS1(fld).Type = adDate Or RS1(fld).Type = adDBDate Or RS1(fld).Type = adDBTimeStamp Then
RS1(fld) = Date
Resume Next
End If
ElseIf err.Number = -2147467259 Then
'fields does not exist in inl and not null
fnm = Trim(Mid(err.Description, 45, 7))
found = 0
For K = 1 To Len(fnm)
If Mid(Trim(fnm), K, 1) = "_" Then
found = 1
Exit For
End If
Next
If found = 1 Then
If RS1(fnm).Type = adChar Or RS1(fnm).Type = adVarChar Or RS1(fnm).Type = adBSTR Or RS1(fnm).Type = adLongVarChar Then
'cn1.Execute "update pnl set " & fnm & "='.'"
RS1(fnm) = ""
Resume Next
ElseIf RS1(fnm).Type = adNumeric Or RS1(fnm).Type = adBigInt Or RS1(fnm).Type = adInteger Or RS1(fnm).Type = adSingle Or RS1(fld).Type = adDouble Then
'cn1.Execute "update pnl set " & fnm & "= 0"
RS1(fnm) = 0
Resume Next
ElseIf RS1(fnm).Type = adDate Or RS1(fnm).Type = adDBDate Or RS1(fnm).Type = adDBTimeStamp Then
'cn1.Execute "update pnl set " & fnm & "=" & Date
RS1(fnm) = Date
Resume Next
End If
End If
ElseIf err.Number = -2147217900 Then
Resume Next
Else
MsgBox err.Description & " " & err.Number & "  " & fld
End If
End Sub


Public Sub prdtoinv()
If rs2.State = 1 Then rs2.Close
    If rs3.State = 1 Then rs3.Close
    rs3.Open "SELECT * FROM PRD", CN1, adOpenDynamic, adLockOptimistic
    While Not rs3.EOF
        CN1.Execute "update inv set sub_unt = " & Val(rs3!sub_unt) & " where prd_idy like '" & Trim(rs3!prd_idy) & "'", d
        CN1.Execute "update inv set prd_uom = " & Val(rs3!con_fac) & " where prd_idy like '" & Trim(rs3!prd_idy) & "'", d
        rs3.MoveNext
    Wend
    CN1.Execute "update inv set stk_typ='Salable'"
    CN1.Execute "update inv set prd_unt=int(prd_qty)"
    CN1.Execute "update inv set prd_pcs=(prd_qty-int(prd_qty))*1000"
    CN1.Execute "update inv set tot_pcs=(prd_qty * sub_unt) + prd_pcs"
    CN1.Execute "update inv set prd_uom=(prd_uom/sub_unt) * tot_pcs"
    If rs2.State = 1 Then rs2.Close
    If rs3.State = 1 Then rs3.Close
End Sub

'Public Function chk1(ss As String)
'Dim k2 As Integer
'k2 = 0
'For k1 = 1 To Len(Trim(ss))
'If Mid(Trim(ss), k1, 1) = "'" Then
'k2 = 1
'Exit For
'End If
'p = p & Mid(Trim(ss), k1, 1)
'Next
'If k2 = 1 Then
'For K = k1 + 1 To Len(Trim(ss))
'p = p & Mid(Trim(ss), K, 1)
'Next
'Else
'p = Trim(ss)
'End If
'chk1 = p
'End Function
