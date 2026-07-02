Attribute VB_Name = "EFF_LIB1"
Option Compare Text
Public VanRef As String
Public MerNum As String

Public osversion As Integer
Public PrdCollection As New Collection
Public EPWD As New encPassword
Public DPWD As New decPassword
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
'Public Declare Function cl_login Lib "NETDUAL.DLL" (ByVal dKey As String) As Integer
'Public Declare Function cl_logout Lib "NETDUAL.DLL" () As Integer
'Public Declare Function cl_lock_ok Lib "NETDUAL.DLL" () As Integer
Public Declare Function cl_login Lib "NETDUAL.DLL" (ByVal dKey As String) As Integer
Public Declare Function cl_get_model Lib "NETDUAL.DLL" (ByRef dBuf As Any) As Integer
Public Declare Function cl_lock_ok Lib "NETDUAL.DLL" () As Integer
Public Declare Function cl_get_id Lib "NETDUAL.DLL" (ByRef dBuf As Any, ByVal dPass As String) As Integer
Public Declare Function cl_get_batch Lib "NETDUAL.DLL" (ByRef dBuf As Any, ByVal dPass As String) As Integer
Public Declare Function cl_set_sign Lib "NETDUAL.DLL" (ByRef dBuf As Any, ByVal dPass As String) As Integer
Public Declare Function cl_get_sign Lib "NETDUAL.DLL" (ByRef dBuf As Any, ByVal dPass As String) As Integer
Public Declare Function cl_set_count Lib "NETDUAL.DLL" (ByVal dCount As Integer, ByVal dPass As String) As Integer
Public Declare Function cl_get_count Lib "NETDUAL.DLL" (ByRef dCount As Integer, ByVal dPass As String) As Integer
Public Declare Function cl_dec_count Lib "NETDUAL.DLL" (ByVal dValue As Integer, ByVal dPass As String) As Integer
Public Declare Function cl_inc_count Lib "NETDUAL.DLL" (ByVal dValue As Integer, ByVal dPass As String) As Integer
Public Declare Function cl_change_pwd Lib "NETDUAL.DLL" (ByVal new_dPass As String, ByVal old_dPass As String) As Integer
Public Declare Function cl_change_wpwd Lib "NETDUAL.DLL" (ByVal new_dwPass As String, ByVal old_dwPass As String) As Integer
Public Declare Function cl_set_osign Lib "NETDUAL.DLL" (ByRef dBuf As Any, ByVal dPass As String) As Integer
Public Declare Function cl_get_osign Lib "NETDUAL.DLL" (ByRef dBuf As Any, ByVal dPass As String) As Integer
Public Declare Function cl_write_block Lib "NETDUAL.DLL" (ByRef dBuf As Any, ByVal dPass As String, ByVal dBlock As Integer, ByVal dPos As Integer, ByVal dCount As Integer) As Integer
Public Declare Function cl_read_block Lib "NETDUAL.DLL" (ByRef dBuf As Any, ByVal dPass As String, ByVal dBlock As Integer, ByVal dPos As Integer, ByVal dCount As Integer) As Integer
Public Declare Function cl_write_word Lib "NETDUAL.DLL" (ByVal dWord As Integer, ByVal dPass As String, ByVal dBlock As Integer, ByVal dPos As Integer) As Integer
Public Declare Function cl_read_word Lib "NETDUAL.DLL" (ByRef dWord As Integer, ByVal dPass As String, ByVal dBlock As Integer, ByVal dPos As Integer) As Integer
Public Declare Function cl_logout Lib "NETDUAL.DLL" () As Integer

' RRDECL32.BAS - VB Function declarations for RRRPT32.DLL.
' Version 8.0 05/27/97

Public Declare Function choosePrinter Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal printer As String, ByVal prSize As Long, ByVal port As String, ByVal poSize As Long) As Long
Public Declare Function chooseReport Lib "rrrpt32.dll" (ByVal appName As String, ByVal library As String, ByVal lsize As Long, ByVal rep As String, ByVal size As Long) As Long
Public Declare Function endReport Lib "rrrpt32.dll" (ByVal hReport As Long) As Long
Public Declare Function endRuntimeInstance Lib "rrrpt32.dll" () As Long
Public Declare Function execRuntime Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bWait As Long, ByVal fsCmdShow As Long, lpCode As Long, lpPageCount As Long, ByVal emsg As String, ByVal emsgsize As Long) As Long
Public Declare Function getAPIVersion Lib "rrrpt32.dll" (lpVersion As Long) As Long
Public Declare Function getBeginPage Lib "rrrpt32.dll" (ByVal hReport As Long, lpPage As Long) As Long
Public Declare Function getCopies Lib "rrrpt32.dll" (ByVal hReport As Long, lpCopies As Long) As Long
Public Declare Function getDBContainer Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal dbContainerPath As String, ByVal size As Long) As Long
Public Declare Function getDisplayErrors Lib "rrrpt32.dll" (ByVal hReport As Long, lpDisperr As Long) As Long
Public Declare Function getDisplayStatus Lib "rrrpt32.dll" (ByVal hReport As Long, lpDispStatus As Long) As Long
Public Declare Function getEndPage Lib "rrrpt32.dll" (ByVal hReport As Long, lpPage As Long) As Long
Public Declare Function getErrorInfo Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal msg As String, ByVal size As Long, lpCode As Long) As Long
Public Declare Function getExportDest Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String) As Long
Public Declare Function getFilter Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String, ByVal size As Long) As Long
Public Declare Function getFilterUsage Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String) As Long
Public Declare Function getFirstFieldName Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal fieldName As String, ByVal size As Long) As Long
Public Declare Function getFirstFilteredFieldName Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal fieldName As String, ByVal size As Long, ByVal filter As Long) As Long
Public Declare Function getFirstGroupField Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal groupFieldName As String, ByVal mSize As Long) As Long
Public Declare Function getFirstRelationInfo Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal relFilePath As String, ByVal relFilePathSize As Long, ByVal relIndexfilePath As String, ByVal relIndexfilePathSizer As Long, ByVal relIndexTag As String, ByVal relIndexTagSize As Long, ByVal relAlias As String, ByVal relAliasSize As Long) As Long
Public Declare Function getFirstSortField Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal sortFieldName As String, ByVal size As Long) As Long
Public Declare Function getFirstUserParam Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal parmname As String, ByVal nSize As Long, ByVal Value As String, ByVal vSize As Long) As Long
Public Declare Function getHighScope Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String, ByVal size As Long) As Long
Public Declare Function getLibrary Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal libname As String, ByVal size As Long) As Long
Public Declare Function getLowScope Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String, ByVal size As Long) As Long
Public Declare Function getMasterIndexInfo Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal masterIndexPath As String, ByVal pathSize As Long, ByVal masterIndexType As String, ByVal masterIndexTag As String, ByVal tagSize As Long) As Long
Public Declare Function getMasterTableName Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal masterTablePath As String, ByVal size As Long) As Long
Public Declare Function getMemoName Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal memoTablePath As String, ByVal size As Long) As Long
Public Declare Function getNewReportHandle Lib "rrrpt32.dll" (ByVal appName As String) As Long
Public Declare Function getNextFieldName Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal fieldName As String, ByVal size As Long) As Long
Public Declare Function getNextFilteredFieldName Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal fieldName As String, ByVal size As Long, ByVal filter As Long) As Long
Public Declare Function getNextGroupField Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal groupFieldName As String, ByVal mSize As Long) As Long
Public Declare Function getNextRelationInfo Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal relFilePath As String, ByVal relFilePathSize As Long, ByVal relIndexfilePath As String, ByVal relIndexfilePathSize As Long, ByVal relIndexTag As String, ByVal relIndexTagSize As Long, ByVal relAlias As String, ByVal relAliasSize As Long) As Long
Public Declare Function getNextSortField Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal sortFieldName As String, ByVal size As Long) As Long
Public Declare Function getNextUserParam Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal parmname As String, ByVal nSize As Long, ByVal Value As String, ByVal vSize As Long) As Long
Public Declare Function getOutputDest Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String) As Long
Public Declare Function getOutputFile Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal outputFilePath As String, ByVal eSize As Long) As Long
Public Declare Function getPreventEscape Lib "rrrpt32.dll" (ByVal hReport As Long, lpNoEsc As Long) As Long
Public Declare Function getPrinter Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal printer As String, ByVal pSize As Long) As Long
Public Declare Function getPrinterPort Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal port As String, ByVal pSize As Long) As Long
Public Declare Function getReportPick Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String) As Long
Public Declare Function getRuntimeRecord Lib "rrrpt32.dll" (ByVal appName As String, ByVal inFile As String) As Long
Public Declare Function getScopeUsage Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String) As Long
Public Declare Function getStatusEveryPage Lib "rrrpt32.dll" (ByVal hReport As Long, lpEveryPage As Long) As Long
Public Declare Function getTestPattern Lib "rrrpt32.dll" (ByVal hReport As Long, lpTest As Long) As Long
Public Declare Function getWinTitle Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal title As String, ByVal tSize As Long) As Long
Public Declare Function initRuntimeInstance Lib "rrrpt32.dll" () As Long
Public Declare Function resetErrorInfo Lib "rrrpt32.dll" (ByVal hReport As Long) As Long
Public Declare Function setBeginPage Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal Page As Long) As Long
Public Declare Function setCopies Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal Copies As Long) As Long
Public Declare Function setDataDir Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal directory As String) As Long
Public Declare Function setDBContainer Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal dbContainerPath As String) As Long
Public Declare Function setDisplayErrors Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bDisperr As Long) As Long
Public Declare Function setDisplayStatus Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bDispStatus As Long) As Long
Public Declare Function setEndPage Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal Page As Long) As Long
Public Declare Function setExportDest Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As Long) As Long
Public Declare Function setFilter Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String) As Long
Public Declare Function setFilterUsage Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As Long) As Long
Public Declare Function setGroupField Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal groupFieldName As String, ByVal groupNum As Long) As Long
Public Declare Function setHighScope Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String) As Long
Public Declare Function setImageDir Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal directory As String) As Long
Public Declare Function setIndexExtension Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal ExtNum As Long) As Long
Public Declare Function setLibrary Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal libname As String) As Long
Public Declare Function setLibraryDir Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal directory As String) As Long
Public Declare Function setLowScope Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String) As Long
Public Declare Function setMasterIndexInfo Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal masterIndexPath As String, ByVal masterIndexType As Long, ByVal masterIndexTag As String) As Long
Public Declare Function setMasterTableName Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal masterTablePath As String) As Long
Public Declare Function setMemoName Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal memoTablePath As String) As Long
Public Declare Function setOutputDest Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As String) As Long
Public Declare Function setOutputFile Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal outputFilePath As String) As Long
Public Declare Function setPreventEscape Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bNoEsc As Long) As Long
Public Declare Function setPrinter Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal printer As String) As Long
Public Declare Function setPrinterPort Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal port As String) As Long
Public Declare Function setRelationInfo Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal relFilePath As String, ByVal relIndexfilePath As String, ByVal relIndexTag As String, ByVal relAlias As String, ByVal aliasNum As Long) As Long
Public Declare Function setReportPick Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As Long) As Long
Public Declare Function setScopeUsage Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal charval As Long) As Long
Public Declare Function setSortField Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal sortFieldName As String, ByVal sortNum As Long) As Long
Public Declare Function setStatusEveryPage Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bEveryPage As Long) As Long
Public Declare Function setStatusFileName Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal fileName As String) As Long
Public Declare Function setSuppressTitle Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bValue As Long) As Long
Public Declare Function setTestPattern Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bTest As Long) As Long
Public Declare Function setUserParam Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal parmname As String, ByVal charval As String) As Long
Public Declare Function setWinBorderStyle Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal eStyle As Long) As Long
Public Declare Function setWinControlBox Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bControlBox As Long) As Long
Public Declare Function setWinHeight Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal iHeight As Long) As Long
Public Declare Function setWinLeft Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal iLeft As Long) As Long
Public Declare Function setWinMaxButton Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bMax As Long) As Long
Public Declare Function setWinMinButton Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bMin As Long) As Long
Public Declare Function setWinParentHandle Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal hParent As Long) As Long
Public Declare Function setWinTitle Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal title As String) As Long
Public Declare Function setWinTop Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal iTop As Long) As Long
Public Declare Function setWinWidth Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal iWidth As Long) As Long
Public Declare Function setWriteAllow Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bWrite As Long) As Long
Public Declare Function setXbaseEditor Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal bXbase As Long) As Long
Public Declare Function writeRuntimeRecord Lib "rrrpt32.dll" (ByVal hReport As Long, ByVal recordFile As String) As Long

' End of R&R rrrpt32.dll functions

Public InputPara1 As String
Public InputPara2 As String
Public InputPara3 As String
Public InputPara4 As String
Public InputPara5 As String
Public ExecuteFunction As Boolean
Public GROUPID, ACCOUNTTYP As String
Public SEL_CMP As String
Public DEL_REC_IDY As String
Public CURDIR As String
Public sprg As String
Public SAL_QRY As String
Public RN As Integer
Public CNU As Integer
Public con As New ADODB.Connection
Public PNV As New ADODB.Connection
Public PNL As New ADODB.Connection
Public RRI As New ADODB.Connection
Public rsrp2 As New ADODB.Recordset
Public INVOICE_DATE As Date
Public LSS_TYP As String
Public FRM_INV_DAT As Date
Public TO_INV_DAT As Date
Public TO_DAT As Date
Public FRM_INV As String
Public END_INV As String
Public rou_idy As String
Public INV_TYP As String
Public PRIDY As String
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
Public cn As New ADODB.Connection
Public CN1 As New ADODB.Connection
Public cn2 As New ADODB.Connection
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public DB As New ADODB.Connection
Public rs As New ADODB.Recordset
Public SRC_RSET As New ADODB.Recordset
Public TGT_RSET As New ADODB.Recordset
Public TB As TableDef
Public DB1 As New ADODB.Connection
Public RS1 As New ADODB.Recordset
Public tb1 As TableDef
Public dbs As New ADODB.Connection
Public TNAME As String
Dim strPath As String
Public Gltxninvid As String
Public Myzip As ColaZip.GSColaZip32
Public Sub AddData(ByRef ctrlName As Control, frmName As Form, fldName As String, tblName As String, Condition As String, Group_order As String)
Dim rs As New ADODB.Recordset
rs.Open "select " & Trim(fldName) & " from " & Trim(tblName) & " " & Condition & " " & Group_order, con, adOpenKeyset, adLockPessimistic
ctrlName.CLEAR
ctrlName.AddItem ""
For A = 1 To rs.RecordCount
    If Trim(rs(0) & "") <> "" Then ctrlName.AddItem Trim(rs(0) & "")
    rs.MoveNext
Next
End Sub
Public Function CHK_IMP_BKP() As Integer
Dim FS As Object, f As Object, F1 As Object, FN As Object
 Dim BKP_INT As Integer
 BKP_INT = 0
 Set FS = CreateObject("Scripting.FileSystemObject")
 Set f = FS.GetFolder(CURDIR)
 Set FN = f.SubFolders
 If FN.count > 0 Then
    For Each F1 In FN
    If Mid(F1, Len(CURDIR) + 2, 3) = "BKP" Then
        If Val(BKP_INT) < Val(Mid(F1, Len(CURDIR) + 2 + 3)) Then
            BKP_INT = Val(Mid(F1, Len(CURDIR) + 2 + 3))
        End If
    End If
    Next
    CHK_IMP_BKP = BKP_INT + 1
Else
    CHK_IMP_BKP = BKP_INT
End If
Set F1 = Nothing
Set f = Nothing
Set FN = Nothing
Set FS = Nothing
End Function
Public Sub SHLPARTYACCOUNT()
If CHECKTB("scl.dbf,snv.dbf,shl.dbf,inv.dbf,inl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
con.Execute "DELETE FROM SCL"
RS1.Open "SELECT * FROM SCL where SCH_IDY = ''", con, adOpenKeyset, adLockPessimistic
rs.Open "SELECT SUM(R_SHL_NOS+DP_SHL_NOS+BRK_SHL),SUM(R_PCK_NOS+DP_PCK_NOS+BRK_PCK),CST_IDY,SHL_IDY,TXN_TYP FROM SNV WHERE TXN_TYP IN('SAL','PRT','STO','AIN') GROUP BY CST_IDY,SHL_IDY", con, adOpenKeyset, adLockPessimistic
While Not rs.EOF = True
    RS1.AddNew
    RS1!INV_DAT = Date
    RS1!Sch_Idy = rs!shl_idy
    RS1!PRD_IDY = rs!CST_IDY
    RS1(3) = 0
    RS1(4) = 0
    RS1(5) = rs(0)
    RS1(6) = rs(1)
    RS1(7) = 0
    RS1(8) = 0
    RS1(9) = 0
    RS1(10) = 0
    RS1(11) = 0
    RS1.Update
    rs.MoveNext
Wend
RS1.Requery
If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(PRD_UNT),SUM(TOT_PCS),SHL_IDY,CST_IDY,INV.TXN_TYP FROM INV,SHL,INL WHERE INL.INV_IDY LIKE INV.INV_IDY AND SHL.PRD_IDY LIKE INV.PRD_IDY AND SHL_TYP='Shell' and INV.STK_TYP IN('Salable','Free') GROUP BY CST_IDY,SHL_IDY,INV.TXN_TYP", con, adOpenKeyset, adLockPessimistic
While Not rs.EOF
    If rs!TXN_TYP = "SAL" Or rs!TXN_TYP = "STO" Or rs!TXN_TYP = "PRT" Then
        con.Execute "update scl set OPN_STK=" & Val(rs(0)) & " where sch_idy like '" & Trim(rs!shl_idy & "") & "' and prd_idy like '" & Trim(rs!CST_IDY) & "'", d
        con.Execute "update scl set RCT_STK=" & Val(rs(1)) & " where sch_idy like '" & Trim(rs!shl_idy & "") & "' and prd_idy like '" & Trim(rs!CST_IDY) & "'", d
        If d = 0 Then
            RS1.AddNew
            RS1!INV_DAT = Date
            RS1!Sch_Idy = rs!shl_idy
            RS1!PRD_IDY = rs!CST_IDY
            RS1(5) = 0
            RS1(6) = 0
            RS1(3) = rs(0)
            RS1(4) = rs(1)
            RS1(7) = 0
            RS1(8) = 0
            RS1(9) = 0
            RS1(10) = 0
            RS1(11) = 0
            RS1.Update
        End If
    ElseIf rs!TXN_TYP = "OPS" Then
    Else
        con.Execute "update scl set ISS_STK=" & Val(rs(0)) & " where sch_idy like '" & Trim(rs!shl_idy & "") & "' and prd_idy like '" & Trim(rs!CST_IDY) & "'", d
        con.Execute "update scl set CLS_STK=" & Val(rs(1)) & " where sch_idy like '" & Trim(rs!shl_idy & "") & "' and prd_idy like '" & Trim(rs!CST_IDY) & "'", d
        If d = 0 Then
            RS1.AddNew
            RS1!INV_DAT = Date
            RS1!Sch_Idy = rs!shl_idy
            RS1!PRD_IDY = rs!CST_IDY
            RS1(3) = 0
            RS1(4) = 0
            RS1(5) = rs(0)
            RS1(6) = rs(1)
            RS1(7) = 0
            RS1(8) = 0
            RS1(9) = 0
            RS1(10) = 0
            RS1(11) = 0
            RS1.Update
        End If
    End If
    A = A + 1
    rs.MoveNext
Wend
If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(R_SHL_NOS+DP_SHL_NOS+BRK_SHL),SUM(R_PCK_NOS+DP_PCK_NOS+BRK_PCK),CST_IDY,SHL_IDY,TXN_TYP FROM SNV WHERE TXN_TYP IN('RCT','PUR') GROUP BY CST_IDY,SHL_IDY", con, adOpenKeyset, adLockPessimistic
While Not rs.EOF
    con.Execute "update scl set OPN_STK=" & Val(rs(0)) & " where sch_idy like '" & Trim(rs!shl_idy & "") & "' and prd_idy like '" & Trim(rs!CST_IDY) & "'", d
    con.Execute "update scl set RCT_STK=" & Val(rs(1)) & " where sch_idy like '" & Trim(rs!shl_idy & "") & "' and prd_idy like '" & Trim(rs!CST_IDY) & "'", d
    If d = 0 Then
        RS1.AddNew
        RS1!INV_DAT = Date
        RS1!Sch_Idy = rs!shl_idy
        RS1!PRD_IDY = rs!CST_IDY
        RS1(3) = rs(0)
        RS1(4) = rs(1)
        RS1(5) = 0
        RS1(6) = 0
        RS1(7) = 0
        RS1(8) = 0
        RS1(9) = 0
        RS1(10) = 0
        RS1(11) = 0
        RS1.Update
    End If
    rs.MoveNext
Wend
con.Execute "update scl set sal_stk= opn_stk-iss_stk"
con.Execute "update scl set fre_stk= rct_stk-cls_stk"
End Sub

Public Function exitform(FRM As Form)
'FRM.Command1.SetFocus
End Function

Public Sub clrctr(FRM As Form)
For A = 1 To FRM.count - 1
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
Public Function convert11(num As String) As String
If Len(Trim(num)) >= 10 Then
    convert11 = Round(num, 2)
Else
    num = Round(Val(num), 2)
    If InStr(num, ".") > 0 Then
        r = Mid(num, InStr(num, ".") + 1, Len(num) - InStr(num, "."))
        If Trim(Len(r)) = 1 Then
            convert11 = Mid(num, 1, InStr(num, ".") - 1) & "." & r & "0"
        ElseIf Trim(Len(r)) = 2 Then
            convert11 = num
        End If
    Else
        If Trim(Len(num)) <> 0 Then
            convert11 = num & ".00"
        Else
            convert11 = "0.00"
        End If
        End If
End If

If Len(convert11) < 10 Then
    Dim i1 As Integer
    i1 = Len(convert11)
    For i = i1 + 1 To 10
    convert11 = " " & convert11
    Next
End If

End Function

Public Function check(tbnme As String, fldnme As String, fldvalue As String) As Boolean
If CHECKTB(tbnme & ".dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from " & tbnme & " where " & fldnme & " like '" & Trim(fldvalue) & "'", con, adOpenStatic
If rs.RecordCount > 0 Then
check = True
Else
check = False
End If
If rs.State = 1 Then rs.Close
End Function
Public Function querycheck(tbnme As String, fldnme As String, fldvalue As String, qry As String) As Boolean
If CHECKTB(tbnme & ".dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from " & tbnme & " where " & fldnme & " like '" & Trim(fldvalue) & "' " & Trim(qry), con, adOpenStatic
If rs.RecordCount > 0 Then
querycheck = True
Else
querycheck = False
End If
If rs.State = 1 Then rs.Close
End Function

Public Function Check1(tbnme As String, fldnme As String, fldvalue As String, TYP As String) As Boolean
If CHECKTB(tbnme & ".dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from " & tbnme & " where " & fldnme & " like '" & Trim(fldvalue) & "' ", con, adOpenStatic
If rs.RecordCount > 0 Then
Check1 = True
Else
Check1 = False
End If
If rs.State = 1 Then rs.Close
End Function

Public Function increment(FLDVAL As String) As String
If CHECKTB("idy.dbf,pre.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If
Dim ict, icr As Integer, icrt As String, rec As New ADODB.Recordset
If rec.State = 1 Then rec.Close
rec.Open "select " & FLDVAL & " from idy", con, adOpenKeyset, adLockPessimistic
If rec.RecordCount = 0 Then
rec.AddNew
rec.Update
End If
rec(0) = Val(rec(0) & "") + 1
ict = Val(rec(0) & "")
rec.Update
If rec.State = 1 Then rec.Close
rec.Open "select pre_" & Mid(FLDVAL, 5, 3) & " from pre", con, adOpenKeyset, adLockPessimistic
If Len(Trim(rec(0) & "") & ict) = 10 Then
    increment = Trim(rec(0) & "") & ict
Else
    icrt = Trim(rec(0) & "")
    Do While Len(icrt) <= 10 - Len(str(ict))
    icrt = icrt & "0"
    Loop
End If
increment = icrt & ict
End Function

Public Function T7increment(FLDVAL As String) As String
If CHECKTB("idy.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If

Dim ict, icr As Integer, icrt As String, rec As New ADODB.Recordset
If rec.State = 1 Then rec.Close
rec.Open "select " & FLDVAL & " from idy", con, adOpenKeyset, adLockPessimistic
If rec.RecordCount = 0 Then
rec.AddNew
rec.Update
End If
rec(0) = Val(rec(0) & "") + 1
ict = Val(rec(0) & "")
rec.Update
Do While Len(icrt) <= 7 - Len(str(ict))
icrt = icrt & "0"
Loop
T7increment = icrt & ict
End Function

Public Function Tincrement(FLDVAL As String) As String
'transactions increment
If Len(FLDVAL) < 7 Then
icrt = ""
Do While Len(icrt) < (7 - Len(Trim(FLDVAL)))
    icrt = icrt & "0"
Loop
Tincrement = Trim(icrt & FLDVAL)
Else
Tincrement = FLDVAL
End If
End Function

Public Function Sincrement(FLDVAL As String) As String
'retreving transactions
Do While Len(icrt) < 10 - Len(FLDVAL)
    icrt = icrt & "0"
Loop
Sincrement = Trim(icrt & FLDVAL)
End Function
Public Function Threement(FLDVAL As String) As String
'retreving transactions
Do While Len(FLDVAL) < 3 - Len(FLDVAL)
    FLDVAL = "0" & FLDVAL
Loop
Threemnt = Trim(FLDVAL)
End Function
Public Function labdisp(FLDVAL As String) As String
If CHECKTB("idy.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select " & FLDVAL & " from idy", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount = 0 Then
rs.AddNew "IDY_OPS", 0
rs.Update
End If
If Val(rs(0) & "") = 0 Then
labdisp = "0000000"
ElseIf Val(rs(0)) > 0 And Val(rs(0)) <= 9 Then
labdisp = "000000" & Val(rs(0))
ElseIf Val(rs(0)) >= 10 And Val(rs(0)) <= 99 Then
labdisp = "00000" & Val(rs(0))
ElseIf Val(rs(0)) >= 100 And Val(rs(0)) <= 999 Then
labdisp = "0000" & Val(rs(0))
ElseIf Val(rs(0)) >= 1000 And Val(rs(0)) <= 9999 Then
labdisp = "000" & Val(rs(0))
ElseIf Val(rs(0)) >= 10000 And Val(rs(0)) <= 99999 Then
labdisp = "00" & Val(rs(0))
ElseIf Val(rs(0)) >= 100000 And Val(rs(0)) <= 999999 Then
labdisp = "0" & Val(rs(0))
ElseIf Val(rs(0)) >= 1000000 And Val(rs(0)) <= 9999999 Then
labdisp = Val(rs(0))
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

Public Function mincrement(NME As String, FLDVAL As String) As String
Do While Len(Trim(icrt)) <= 4 - Len(Trim(FLDVAL))
icrt = icrt & "0"
Loop
mincrement = icrt & FLDVAL
End Function
Public Sub Check_Academic()
Dim CL_ACA As Integer
CL_ACA = cl_lock_ok
If CL_ACA <> 1 Then
   MsgBox "DAS Version 4.0 License Hardware Lock (Dongle) Missing! Check Device Connection to Parallel Port - USB!", vbCritical, "DAS Version 4.0"
   End
End If
Dim dBuf(50) As Byte
dSign = ""
dResult = cl_get_osign(dBuf(0), "")
If dResult = 1 Then
      For i = 0 To 49
        dSign = dSign + Chr(dBuf(i))
      Next
End If
Dim CL_GML As String
Dim mBuf(50) As Byte
dModel = cl_get_model(mBuf(0))
If dModel = 1 Then
For i = 0 To 49
    CL_GML = CL_GML + Chr(mBuf(i))
Next
End If
Dim Dng_Stmp As New ADODB.Recordset
If Trim(Left(dSign, 8)) = "DASHCCBL" Then
    'MsgBox ("DAS Academic Version, HCCBPL Asset"), vbInformation, "DAS Version 4.0"
    M_MDI.Caption = "DAS Version 4.01.01 Rel. - Distributor Automation System, Ver.Rel.04-Sep-2009 " & "HCCBPL Asset, Academic License " & "Login : " & Trim(TXTUSERNAME) & " for " & Trim(SEL_CMP)
    'M_MDI.Caption = "DAS Version 4.0 - Distributor Automation System" & Space(60) & "HCCBPL Asset, Academic License LOGIN : " & Trim(TXTUSERNAME) & " for " & Trim(SEL_CMP)
    If Dng_Stmp.State = 1 Then Dng_Stmp.Close
    Dng_Stmp.Open "SELECT * FROM DIR WHERE DIR_IDY = 'DNG' .AND. DIR_IDX = 'STM'", con, adOpenKeyset, adLockPessimistic
    If Dng_Stmp.RecordCount > 0 Then
        Dng_Stmp("DIR_ALS") = "Academic"
        Dng_Stmp("DIR_DAT") = Format(Date, "MM/DD/YYYY")
        Dng_Stmp("DIR_MSG") = Trim(CL_GML)
        Dng_Stmp.Update
    Else
        Dng_Stmp.AddNew
        Dng_Stmp("DIR_IDY") = "DNG"
        Dng_Stmp("DIR_IDX") = "STM"
        Dng_Stmp("DIR_ALS") = "Academic"
        Dng_Stmp("DIR_MSG") = Trim(CL_GML)
        Dng_Stmp("DIR_DAT") = Format(Date, "MM/DD/YYYY")
        Dng_Stmp.Update
    End If
Else
    If Dng_Stmp.State = 1 Then Dng_Stmp.Close
    Dng_Stmp.Open "SELECT * FROM DIR WHERE DIR_IDY = 'DNG' .AND. DIR_IDX = 'STM'", con, adOpenKeyset, adLockPessimistic
    If Dng_Stmp.RecordCount > 0 Then
        Dng_Stmp("DIR_ALS") = "Regular"
        Dng_Stmp("DIR_DAT") = Format(Date, "MM/DD/YYYY")
        Dng_Stmp("DIR_MSG") = Trim(CL_GML)
        Dng_Stmp.Update
    Else
        Dng_Stmp.AddNew
        Dng_Stmp("DIR_IDY") = "DNG"
        Dng_Stmp("DIR_IDX") = "STM"
        Dng_Stmp("DIR_ALS") = "Regular"
        Dng_Stmp("DIR_DAT") = Format(Date, "MM/DD/YYYY")
        Dng_Stmp("DIR_MSG") = Trim(CL_GML)
        Dng_Stmp.Update
    End If
End If
Set Dng_Stmp = Nothing
End Sub
Public Sub checkdongle()
Dim CL_INT As Integer
CL_INT = cl_lock_ok
If CL_INT <> 1 Then
   MsgBox "DAS Version 4.0 License Hardware Lock (Dongle) Missing! Check Device Connection to Parallel Port - USB!", vbCritical, "DAS Version 4.0"
   End
End If
End Sub

Public Sub LOGIN()
If cl_login("067670508880070277000726") <> 1 Then
    MsgBox "DAS Version 4.0 License Hardware Lock (Dongle) Missing! Check Device Connection to Parallel Port - USB!", vbCritical, "DAS Version 4.0"
End
End If
End Sub
Public Sub LOGOUT()
If CHECKTB("log.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

Dim RSlog As New ADODB.Recordset
RSlog.Open "select * from log", con, adOpenKeyset, adLockPessimistic
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
If CHECKTB("yer.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'FTD'") = True Then
    M_ADD_DYS = 0#
    M_ADD_DYS = GetValue("DIR_ALS", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'FTD'")
    If Val(M_ADD_DYS) = 0 Then
        M_ADD_DYS = 0
    Else
        M_ADD_DYS = Val(M_ADD_DYS)
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM YER WHERE ACT_PAS LIKE 'Y'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then rs.MoveFirst
    For A = 0 To FRM.count - 1
        If TypeOf FRM.Controls(A) Is DTPicker Then
            'FRM.Controls(A).MaxDate = rs("TO_DAT")
            SLD = Format(DateAdd("D", Val(M_ADD_DYS), Date), "MM/DD/YYYY")
            If SLD <= Trim(rs("TO_DAT")) Then
                FRM.Controls(A).MaxDate = SLD
            Else
                FRM.Controls(A).MaxDate = Trim(rs("TO_DAT"))
            End If
            FRM.Controls(A).MinDate = Trim(rs("FRM_DAT"))
            FRM.Controls(A).Value = Date
            con.Execute "UPDATE DIR SET DIR_ALS = '" & Format(rs("FRM_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
            con.Execute "UPDATE DIR SET DIR_MSG = '" & Format(rs("TO_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
        End If
    Next
    Exit Sub
Else
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM YER WHERE ACT_PAS LIKE 'Y'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then rs.MoveFirst
    For A = 0 To FRM.count - 1
        If TypeOf FRM.Controls(A) Is DTPicker Then
            FRM.Controls(A).MaxDate = Trim(rs("TO_DAT"))
            FRM.Controls(A).MinDate = Trim(rs("FRM_DAT"))
            FRM.Controls(A).Value = Date
            con.Execute "UPDATE DIR SET DIR_ALS = '" & Format(rs("FRM_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
            con.Execute "UPDATE DIR SET DIR_MSG = '" & Format(rs("TO_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
        End If
    Next
    Exit Sub
End If
Label:
MsgBox "Invalid Financial Year & Accounting Period, Contact DAS Version 4.0 Support!", vbCritical, "DAS Version 4.0"
End
End Sub

Public Sub SETYEARNEW(FRM As Form)
On Error GoTo Label
If CHECKTB("yer.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM YER WHERE ACT_PAS LIKE 'Y'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs.MoveFirst
For A = 0 To FRM.count - 1
    If TypeOf FRM.Controls(A) Is DTPicker Then
        FRM.Controls(A).MaxDate = rs("FRM_DAT")
        FRM.Controls(A).MinDate = rs("FRM_DAT")
        FRM.Controls(A).Value = rs("FRM_DAT")
        con.Execute "UPDATE DIR SET DIR_ALS = '" & Format(rs("FRM_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
        con.Execute "UPDATE DIR SET DIR_MSG = '" & Format(rs("TO_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
    End If
Next
Exit Sub
Label:
MsgBox "Invalid Accounting Period, Contact DAS Version 4.0 Support!", vbCritical, "DAS Version 4.0"
End
End Sub
Public Sub SETYEAROPEN(FRM As Form)
On Error GoTo Label
If CHECKTB("yer.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If
If FRM.Caption = "DB Quick Restore" Then
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'AYR'") = True Then
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM YER WHERE ACT_PAS LIKE 'Y'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then rs.MoveFirst
    For A = 0 To FRM.count - 1
        If TypeOf FRM.Controls(A) Is DTPicker Then
            FRM.Controls(A).MaxDate = Format("31/03/2015", "DD/MM/YYYY")
            FRM.Controls(A).MinDate = Format("01/04/1999", "DD/MM/YYYY")
            FRM.Controls(A).Value = Date$
            con.Execute "UPDATE DIR SET DIR_ALS = '" & Format(rs("FRM_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
            con.Execute "UPDATE DIR SET DIR_MSG = '" & Format(rs("TO_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
        End If
    Next
    Exit Sub
Else
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM YER WHERE ACT_PAS LIKE 'Y'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then rs.MoveFirst
    For A = 0 To FRM.count - 1
        If TypeOf FRM.Controls(A) Is DTPicker Then
            FRM.Controls(A).MaxDate = rs("TO_DAT")
            FRM.Controls(A).MinDate = rs("FRM_DAT")
            FRM.Controls(A).Value = Date$
            con.Execute "UPDATE DIR SET DIR_ALS = '" & Format(rs("FRM_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
            con.Execute "UPDATE DIR SET DIR_MSG = '" & Format(rs("TO_DAT"), "DD/MM/YYYY") & "' WHERE DIR_IDY LIKE 'CUR'"
        End If
    Next
    Exit Sub
End If
Label:
MsgBox "Invalid Financial Year & Accounting Period, Contact DAS Version 4.0 Support!", vbCritical, "DAS Version 4.0"
End
End Sub
Public Function fragment(amt As String) As String
If Len(Trim(amt)) = 10 Then
    fragment = Mid(amt, 1, 2) & "," & Mid(amt, 3, 2) & "," & Mid(amt, Len(amt) - 6 + 1, 6)
ElseIf Len(Trim(amt)) = 9 Then
    fragment = Mid(amt, 1, 1) & "," & Mid(amt, 2, 2) & "," & Mid(amt, Len(amt) - 6 + 1, 6)
ElseIf Len(Trim(amt)) = 8 Then
    fragment = Mid(amt, 1, 2) & "," & Mid(amt, Len(amt) - 6 + 1, 6)
ElseIf Len(Trim(amt)) = 7 Then
    fragment = Mid(amt, 1, 1) & "," & Mid(amt, Len(amt) - 6 + 1, 6)
Else
    fragment = amt
End If
End Function

Public Function CONQTY(ByVal EQTY As Double, PRDID As String) As Double
If CHECKTB("prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in "
    Exit Function
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select sub_unt,cur_lev from prd where prd_idy like '" & Trim(PRDID) & "'", con, adOpenKeyset, adLockPessimistic
ENT_QTY = MTS(Val(EQTY), rs(0))
CUR_QTY = MTS(Val(rs(1)), rs(0))
CONQTY = STM(Val(CUR_QTY) - Val(ENT_QTY), rs(0))
End Function

Public Function CONPRDQTY(ByVal CURQTY As String, ByVal EQTY As String, PRDID As String) As Long
If CHECKTB("prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
EQTY = QTY(EQTY)
CURQTY = QTY(CURQTY)
rs.Open "select sub_unt from prd where prd_idy like '" & PRDID & "'", con, adOpenKeyset, adLockPessimistic
If Val(rs(0)) > 0 Then
    A = InStr(EQTY, ".")
    b = InStr(CURQTY, ".")
    If b > 0 Then
        CDECNUM = Mid(CURQTY, b + 1, 3)
        NUM2 = Mid(CURQTY, 1, b - 1)
        TOT1 = Val(rs(0) * NUM2) + Val(CDECNUM)
    End If
    
    If A > 0 Then
    EDECNUM = Mid(EQTY, A + 1, 3)
    NUM1 = Mid(EQTY, 1, A - 1)
    TOT2 = Val(rs(0) * Val(NUM1)) + Val(EDECNUM)
    End If
    CONPRDQTY = TOT1 - TOT2
Else
    CONPRDQTY = QTY(Round(CURQTY - EQTY, 3))
End If
End Function

Public Function ADDQTY(ByVal EQTY As Double, PRDID As String) As Double
If CHECKTB("prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select sub_unt,cur_lev from prd where PRD_idy like '" & Trim(PRDID) & "'", con, adOpenKeyset, adLockPessimistic
ENT_QTY = MTS(Val(EQTY), rs(0))
CUR_QTY = MTS(Val(rs(1)), rs(0))
ADDQTY = STM(Val(ENT_QTY) + Val(CUR_QTY), rs(0))
End Function

Public Function ADDBRDQTY(ByVal EQTY As Double, BRDID As String) As Double
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select sub_unt,prd_qty from prd,brd where prd.prd_idy like brd.PRD_idy and BRD.bat_idy like '" & Trim(BRDID) & "'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then                  'Record EOF status check incorporated by PNR here 20.06.2009
    rs.MoveFirst
    ENT_QTY = MTS(Val(EQTY), rs(0))
    CUR_QTY = MTS(Val(rs(1) & ""), rs(0))
    ADDBRDQTY = STM(Val(ENT_QTY) + Val(CUR_QTY), rs(0))
Else
    ENT_QTY = 0
    CUR_QTY = 0
    ADDBRDQTY = 0
End If
End Function
', Optional DEFA As String
Public Function CONBRDQTY(ByVal EQTY As Double, BRDID As String) As Double
If CHECKTB("prd.dbf,brd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select sub_unt,prd_qty from prd,brd where prd.prd_idy like brd.PRD_idy and bat_idy like '" & Trim(BRDID) & "'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    ENT_QTY = MTS(Val(EQTY), rs(0))
    CUR_QTY = MTS(Val(rs(1)), rs(0))
    CONBRDQTY = STM(Val(CUR_QTY) - Val(ENT_QTY), rs(0))
Else
    ENT_QTY = 0
    CUR_QTY = 0
End If
'If Trim(DEFA) = "+" Then
'    CONBRDQTY = STM(Val(CUR_QTY) + Val(Ent_qty), RS(0))
'ElseIf Trim(DEFA) = "-" Then
    
'End If
End Function

Public Function VANADDQTY(EQTY As String, PRDID As String, invid As String) As String
If CHECKTB("prd.dbf,inv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If

Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
EQTY = QTY(EQTY)

rs.Open "select sub_unt from prd where prd_idy like '" & Trim(PRDID) & "'", con, adOpenKeyset, adLockPessimistic
RS1.Open "select RET_QTY from inv where prd_idy like '" & Trim(PRDID) & "' and txn_typ like 'RSI' AND inv_idy like '" & Trim(invid) & "'", con, adOpenKeyset, adLockPessimistic
If Val(rs(0)) > 0 Then
    A = InStr(QTY(EQTY), ".")
    b = InStr(QTY(RS1(0) & ""), ".")
    
    If A > 0 Then
    EDECNUM = Mid(EQTY, A + 1, 3)
    NUM1 = Mid(EQTY, 1, A - 1)
    TOT1 = Val(Val(rs(0) & "") * Val(NUM1)) + Val(EDECNUM)
    End If
    If b > 0 Then
    CDECNUM = Mid(QTY(RS1(0) & ""), b + 1, 3)
    NUM2 = Mid(QTY(RS1(0) & ""), 1, b - 1)
    TOT2 = Val(Val(rs(0)) * Val(NUM2)) + Val(CDECNUM)
    End If
    'tot2 = Val(rs(0) * num2) + Round(Val(cdecnum), 3)
    tot = TOT1 + TOT2
    r = Round(Val(tot) Mod Val(rs(0)), 3)       'remainder
    V = Val(tot) \ Val(rs(0))   'quotient
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

Public Function VANCONQTY(EQTY As String, PRDID As String, invid As String) As String
If CHECKTB("prd.dbf,inv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If

Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
EQTY = QTY(EQTY)

rs.Open "select sub_unt from prd where prd_idy like '" & Trim(PRDID) & "'", con, adOpenKeyset, adLockPessimistic
RS1.Open "select RET_qty from inv where prd_idy like '" & Trim(PRDID) & "' and txn_typ like 'RSI' AND inv_idy like '" & Trim(invid) & "'", con, adOpenKeyset, adLockPessimistic
If Val(rs(0)) > 0 Then
    A = InStr(QTY(EQTY), ".")
    b = InStr(QTY(RS1(0) & ""), ".")
    
    If A > 0 Then
    EDECNUM = Mid(EQTY, A + 1, 3)
    NUM1 = Mid(EQTY, 1, A - 1)
    TOT1 = Val(rs(0) * NUM1) + Val(EDECNUM)
    End If
       
    If b > 0 Then
    CDECNUM = Mid(QTY(RS1(0) & ""), b + 1, 3)
    NUM2 = Mid(QTY(RS1(0) & ""), 1, b - 1)
    TOT2 = (Val(rs(0) & "") * Val(NUM2)) + Val(CDECNUM)
    End If
    'tot2 = Val(rs(0) * num2) + Round(Val(cdecnum), 3)
    tot = TOT1 - TOT2
    
    r = Round(Val(tot) Mod Val(rs(0)), 3)       'remainder
    V = Val(tot) \ Val(rs(0))   'quotient
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

Public Function RATCAL(unt As String, RAT As String, SUBUNT As String) As String
unt = QTY(unt)
A = InStr(QTY(unt), ".")
If A > 0 Then
    EDECNUM = Mid(unt, A + 1, 3)
    NUM1 = Mid(unt, 1, A - 1)
End If
C = (Val(EDECNUM) * Val(RAT))
If C <> 0 Then
V = C / Val(SUBUNT)
Else
V = 0
End If
RATCAL = convert(amt((Val(NUM1) * Val(RAT)) + Val(V)))
End Function

Public Function WGTCAL(unt As String, CONFAC As String, SUBUNT As String) As String
unt = QTY(unt)
A = InStr(QTY(unt), ".")
If A > 0 Then
    EDECNUM = Mid(unt, A + 1, 3)
    NUM1 = Mid(unt, 1, A - 1)
End If
WGTCAL = QTY((Val(NUM1) * Val(CONFAC)) + Round(((Val(EDECNUM) * Val(CONFAC)) / Val(SUBUNT)), 3))
End Function

Public Function UNTCAL(unt As String, PRDID As String) As String
If CHECKTB("prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If

Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM PRD WHERE PRD_IDY LIKE '" & Trim(PRDID) & "'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
A = Int(Val(unt) / Val(rs("CON_FAC") & ""))
C = Round((Val(rs("CON_FAC") * 1000) / Val(rs("SUB_UNT") * 1000)), 3)
b = Val(unt) * 1000 Mod C * 1000
If b <> 0 Then MsgBox "Conversion Factor and Sub Units Mismatch!", vbExclamation
b = Val(unt) * 1000 Mod Val(rs("CON_FAC") & "") * 1000
b = b / (Val(rs("CON_FAC") & "") / Val(rs("SUB_UNT") & "") * 1000)
b = b / 1000
UNTCAL = QTY(Val(A) + Val(b))
Else
UNTCAL = 0
End If
End Function

Public Sub TRAIL_BALANCE(TO_DAT1 As Date)
TO_DAT = TO_DAT1
Unload M_FRM_ASD
M_FRM_APD.Tag = "TB"
M_FRM_APD.Show 1
End Sub

Public Sub PAL()
If CHECKTB("opa.dbf,act.dbf,typ.dbf,dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

Dim MLOS As Long, MPRO As Long
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
Dim RS1 As New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
Dim rs2 As New ADODB.Recordset
If rs2.State = 1 Then rs2.Close
rs.Open "select * from opa", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    If RS1.State = 1 Then RS1.Close
    RS1.Open "select * from act WHERE act_idy like '" & Trim(rs("act_idy") & "") & "'", con, adOpenKeyset, adLockPessimistic
    If RS1.RecordCount > 0 Then
        If rs2.State = 1 Then rs2.Close
        rs2.Open "select * from typ where typ_def like '" & Trim(RS1("grp_clf") & "") & "'", con, adOpenKeyset, adLockPessimistic
        If rs2.RecordCount > 0 Then
            If Trim(rs2("typ_fac") & "") = "D" And Trim(rs2("typ_clf") & "") = "PL" Then
                MLOS = MLOS + Val(rs("opa_dbt"))
            ElseIf Trim(rs2("typ_fac") & "") = "C" And Trim(rs2("typ_clf") & "") = "PL" Then
                MPRO = MPRO + Val(rs("opa_crt"))
            End If
        End If
    End If
    rs.MoveNext
Next
con.Execute "update dir set dir_als='" & MLOS & "', dir_msg='" & MPRO & "' where dir_idy like 'PRO' and dir_idx like 'LOS'"
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
    C = Val(b) + (M_MOD / 1000)
Else
    C = M_QTY
End If
If M_CHK = True Then C = C * -1
P_QTY_UNT = QTY(Trim(C))
End Function
Public Sub M_HELP()
On Error GoTo LBL
M_PATH = App.Path & "\CONTHELP.CHM"
If Company.SysInfo1.OSPlatform = 1 Then
    Shell "C:\WINDOWS\HH.EXE " & M_PATH, vbMaximizedFocus
ElseIf Company.SysInfo1.OSPlatform = 2 Then
    Shell "C:\WINNT\HH.EXE " & M_PATH, vbMaximizedFocus
Else
    Shell "C:\WINDOWS\HH.EXE " & M_PATH, vbMaximizedFocus
End If
Exit Sub
LBL:
'MsgBox err.Description, vbCritical
MsgBox "Help will be made available shortly!", vbInformation, "DAS Version 4.0"
End Sub
Public Sub M_HELP_REG()
On Error GoTo LBL
M_PATH = App.Path & "\CONTHELP.CHM"
If Company.SysInfo1.OSPlatform = 1 Then
    Shell "C:\WINDOWS\HH.EXE " & M_PATH, vbMaximizedFocus
ElseIf Company.SysInfo1.OSPlatform = 2 Then
    Shell "C:\WINNT\HH.EXE " & M_PATH, vbMaximizedFocus
Else
    Shell "C:\WINDOWS\HH.EXE " & M_PATH, vbMaximizedFocus
End If
Exit Sub
LBL:
'MsgBox err.Description, vbCritical
MsgBox "Help will be made available shortly!", vbInformation, "DAS Version 4.0"
End Sub
Public Function OpenWinReport(ReportLibrary As String, ReportName As String, Query As Integer, filter As String, Destination As Integer)
If Dir(Trim(ReportLibrary)) = "" Then
    MsgBox "Report :" & Trim(ReportName) & " - " & "Library " & Trim(ReportLibrary) & " Not Found! Contact Help Desk!", vbInformation, "DAS Version 4.0"
    Exit Function
End If
With M_MDI.RRReport1
    .ReportLibrary = Trim(ReportLibrary)
    .ReportName = Trim(ReportName)
    .ResetProperties = True
    .DisplayError = True
    .ReportDirectory = "C:\EWinReports"
    .printer = "?"
    .Query = Val(Query)
    .filter = filter
    .Destination = Val(Destination)
    .RunReport 1
End With
M_MDI.RRReport1.ResetProperties = True
M_MDI.RRReport1.ResetControl
End Function
Public Function OpenReport(ReportLibrary As String, ReportName As String, Query As Integer, filter As String, Destination As Integer)
'endRuntimeInstance
On Error GoTo RPT_ERROR
If Dir(Trim(ReportLibrary)) = "" Then
    MsgBox "Report :" & Trim(ReportName) & " - " & "Library " & Trim(ReportLibrary) & " Not Found! Contact Help Desk!", vbInformation, "DAS Version 4.0"
    Exit Function
End If

                    ' Code Under Consideration for inc. by PNR
                    '''''Dim LPT_CHK As New ADODB.Recordset
                    '''''If LPT_CHK.State = 1 Then LPT_CHK.Close
                    '''''LPT_CHK.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'LPT' AND DIR_IDX LIKE 'YES'", CON, adOpenKeyset, adLockPessimistic
                    '''''If LPT_CHK.EOF = False Then
                    '''''    If Trim(LPT_CHK("DIR_MSG")) = "Y" Then
                    '''''        With M_MDI.RRReport1
                    '''''            .ReportLibrary = Trim(ReportLibrary)
                    '''''            .ReportName = Trim(ReportName)
                    '''''            .ResetProperties = True
                    '''''            .DisplayError = True
                    '''''            .Query = Val(Query)
                    '''''            .Filter = Filter
                    '''''            .ExportDestination = 2
                    '''''            .PrintFileName = "C:\DEX\LATREP.TXT"
                    '''''            .RunReport 1
                    '''''        End With
                    '''''        If MsgBox("Print to LPT1 DotMatrix Printer?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                    '''''            CopyFile "C:\DEX\LATREP.TXT", "PRN", A
                    '''''        End If
                    '''''    End If
                    '''''Else
                    '''''    LPT_CHK.AddNew
                    '''''    LPT_CHK("DIR_IDY") = "LPT"
                    '''''    LPT_CHK("DIR_IDX") = "YES"
                    '''''    LPT_CHK("DIR_MSG") = "N"
                    '''''    LPT_CHK("DIR_ALS") = ""
                    '''''    LPT_CHK("DIR_DAT") = Date
                    '''''    LPT_CHK.Update
                    '''''End If
                    '''''LPT_CHK.Close

With M_MDI.RRReport1
    .IndexExtension = 1
    .WindowBorderStyle = True
    .ReportLibrary = Trim(ReportLibrary)
    .ReportName = Trim(ReportName)
    .ResetProperties = True
    .DisplayError = True
    .printer = "?"
    .Query = Val(Query)
    .filter = filter
    .Destination = Val(Destination)
    .UpdateControl = True
    .RunReport 1
End With
M_MDI.RRReport1.ResetProperties = True
M_MDI.RRReport1.ResetControl
Exit Function
RPT_ERROR:
MsgBox err.Description, vbCritical, "DAS Version 4.0"
M_MDI.RRReport1.ResetProperties = True
M_MDI.RRReport1.ResetControl
End Function
Public Sub OpenScopeReport(RP1 As String, REPNAME As String, qry As Integer, SCP As Integer, LOW As String, HIGH As String, filter As String, Destination As Integer)
If Dir(Trim(ReportLibrary)) = "" Then
    MsgBox "Report :" & Trim(REPNAME) & " - " & "Library " & Trim(RP1) & " Not Found! Contact Help Desk!", vbInformation, "DAS Version 4.0"
    Exit Sub
End If
With M_MDI.RRReport1
    .ReportLibrary = RP1
    .ReportName = REPNAME
    .Query = qry
    .scope = SCP
    .LowScope = LOW
    .HighScope = HIGH
    .filter = filter
    .Destination = Destination
    .printer = "?"
    .RunReport 1
End With
M_MDI.RRReport1.ResetProperties = True
M_MDI.RRReport1.ResetControl
End Sub
Public Function OpenDocument(ReportLibrary As String, ReportName As String, scope As Integer, loscope As String, hiscope As String, Destination As Integer, mindex As String)
On Error GoTo ERR_TRAP
If CHECKTB("dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Function
End If

Dim rsdir As New ADODB.Recordset
With M_MDI.RRReport1
    .ReportLibrary = Trim(ReportLibrary)
    If ReportName = "Invoice" Then
        rsdir.Open "select * from dir where dir_idy like 'AUT' and dir_idx like 'IND'", con, adOpenKeyset, adLockPessimistic
        If rsdir.RecordCount > 0 Then
            If Trim(rsdir("DIR_MSG")) = "" Then
                .ReportName = "Invoice"
            Else
                .ReportName = Trim(rsdir("DIR_MSG"))
            End If
        End If
    Else
        .ReportName = Trim(ReportName)
    End If
    .printer = "?"
    .scope = Val(scope)
    .MasterIndex = Trim(mindex)
    .LowScope = loscope
    .HighScope = hiscope
    .Destination = Val(Destination)
    .RunReport 1
End With
Exit Function
ERR_TRAP:
MsgBox err.Description
End Function
Public Sub RRW_OpenDocument(ReportLibrary As String, ReportName As String, scope As Integer, loscope As String, hiscope As String, Destination As Integer, mindex As String)
If CHECKTB("dir.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If
Dim rsdir As New ADODB.Recordset
rsdir.Open "select * from dir where dir_idy like 'AUT' and dir_idx like 'IND'", con, adOpenKeyset, adLockPessimistic
If rsdir.RecordCount > 0 Then
    If Trim(rsdir("DIR_MSG")) = "" Then
        ReportName = "INVOICE.RRW"
    Else
        ReportName = Trim(rsdir("DIR_MSG")) & ".RRW"
    End If
End If
If Dir(Trim(ReportName)) = "" Then
    MsgBox "Report :" & Trim(ReportName) & " Not Found in the Folder! " & CURDIR & " Contact Help Desk!", vbInformation, "DAS Version 4.0"
    Exit Sub
End If
On Error GoTo ERR_TRAP
With M_MDI.RRReport1
'    .ReportLibrary = Trim(ReportLibrary)
    .ReportName = Trim(ReportName)
    .printer = "?"
    .scope = Val(scope)
    .MasterIndex = Trim(mindex)
    .LowScope = loscope
    .HighScope = hiscope
    .Destination = Val(Destination)
    .RunReport 1
End With
Exit Sub
ERR_TRAP:
MsgBox err.Description
End Sub

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
If CHECKTB("prd.dbf,inv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

If MsgBox("Update Stocks..", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
    Dim REC_SET As New ADODB.Recordset
    Dim REC_SET1 As New ADODB.Recordset
    
    con.Execute "UPDATE PRD SET CUR_LEV = 0"
    REC_SET.Open "SELECT * FROM INV where txn_typ in('OPS','PUR','SRN','STO','RSR','PCR')", con, adOpenKeyset, adLockPessimistic
    If REC_SET.RecordCount > 0 Then
       For A = 1 To REC_SET.RecordCount
                If REC_SET1.State = 1 Then REC_SET1.Close
                REC_SET1.Open "SELECT * FROM PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET("PRD_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If REC_SET1.RecordCount > 0 Then
                    PRD_SUB = P_SUB_UNT(QTY(REC_SET1("CUR_LEV") & ""), REC_SET1("SUB_UNT"))
                    INV_SUB = P_SUB_UNT(QTY(REC_SET("PRD_QTY") & ""), REC_SET1("SUB_UNT"))
                    REC_SET1("CUR_LEV") = P_QTY_UNT(Val(PRD_SUB) + Val(INV_SUB), REC_SET1("SUB_UNT"))
                    REC_SET1.Update
                End If
                REC_SET.MoveNext
        Next
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM INV where txn_typ in('LSL','SAL','PRT','STI','RSI','PCS')", con, adOpenKeyset, adLockPessimistic
        For A = 1 To REC_SET.RecordCount
            If REC_SET1.State = 1 Then REC_SET1.Close
            REC_SET1.Open "SELECT * FROM PRD WHERE PRD_IDY LIKE '" & Trim(REC_SET("PRD_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
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

Public Sub STL_LED(FRM_DAT As Date, TO_DAT As Date, qry As String, T As String)
If CHECKTB("stl.dbf,inv.dbf,inl.dbf,prd.dbf,pgr.dbf,cld.dbf,brd.dbf,ops.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

M_FRM_PPQ.MousePointer = vbHourglass
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
con.Execute "delete from stl"
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
rs.Open "select * from stl", con, adOpenKeyset, adLockPessimistic
'RS1.Open "select * from inl,INV,PRD WHERE STK_UPD ='Y' AND PRD.PRD_IDY LIKE INV.PRD_IDY AND inv.inv_idy like inl.inv_idy and INV.TXN_TYP IN ('PUR','SRN','STI','DMR','OPS')  and INV.inv_dat >= ctod('" & FRM_DAT & "') AND INV.INV_DAT <=CTOD('" & TO_DAT & "')" & Trim(qry), CON, adOpenKeyset, adlockpessimistic
RS1.Open "select * from inl,INV,PRD WHERE PRD.PRD_IDY LIKE INV.PRD_IDY AND inv.inv_idy like inl.inv_idy and INV.TXN_TYP IN ('PUR','SRN','STI','DMR','OPS')  and INV.inv_dat >= ctod('" & FRM_DAT & "') AND INV.INV_DAT <=CTOD('" & TO_DAT & "')" & Trim(qry), con, adOpenKeyset, adLockPessimistic 'Modified by Radhika & Ramanesh
For A = 1 To RS1.RecordCount
    rs.AddNew
    rs("inv_idy") = Trim(RS1("inv_idy"))
    rs("inv_dat") = Trim(RS1("inv_dat"))
    rs("prd_idy") = Trim(RS1("prd_idy"))
    rs("ISS_qty") = 0
    rs("RCT_qty") = Trim(RS1("prd_qty"))
    rs("txn_typ") = Trim(RS1("txn_typ"))
    rs("cst_idy") = Trim(RS1("cst_idy"))
    rs("TXN_RMK") = Trim(RS1("TXN_RMK") & "")
    rs.Update
    RS1.MoveNext
Next
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
rs.Open "select * from stl", con, adOpenKeyset, adLockPessimistic
'RS1.Open "select inv.* from INV,PRD,PGR WHERE STK_UPD ='Y' AND PRD.PRD_IDY LIKE INV.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND inv.TXN_TYP IN ('SAL','PRT','STO','ADJ') AND inv.inv_dat >= ctod('" & FRM_DAT & "') AND INV.INV_DAT <=CTOD('" & TO_DAT & "')" & Trim(qry), CON, adOpenKeyset, adlockpessimistic
RS1.Open "select inv.* from INV,PRD,PGR WHERE PRD.PRD_IDY LIKE INV.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND inv.TXN_TYP IN ('SAL','PRT','STO','ADJ') AND inv.inv_dat >= ctod('" & FRM_DAT & "') AND INV.INV_DAT <=CTOD('" & TO_DAT & "')" & Trim(qry), con, adOpenKeyset, adLockPessimistic ' Modified by Radhika & Ramanesh
For A = 1 To RS1.RecordCount
    rs.AddNew
    rs("inv_idy") = Trim(RS1("inv_idy"))
    rs("inv_dat") = Trim(RS1("inv_dat"))
    rs("prd_idy") = Trim(RS1("prd_idy"))
    If Trim(RS1("stk_typ")) = "Salable" Or Trim(RS1("stk_typ")) = "Free" Then
        rs("rct_qty") = 0
        rs("iss_qty") = Trim(RS1("prd_qty"))
    ElseIf Trim(RS1("stk_typ")) = "TakeBack" Or Trim(RS1("stk_typ")) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage" Then
        rs("iss_qty") = 0
        rs("rct_qty") = Trim(RS1("prd_qty"))
    End If
    rs("txn_typ") = Trim(RS1("txn_typ"))
    If rs("txn_typ") = "FRS" Or rs("TXN_TYP") = "FRC" Then
        If rs2.State = 1 Then rs2.Close
        rs2.Open "SELECT CST_IDY,TXN_RMK FROM INL WHERE INV_IDY LIKE '" & Trim(RS1("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
    Else
        If rs2.State = 1 Then rs2.Close
        rs2.Open "SELECT CST_IDY,TXN_RMK FROM INL WHERE INV_IDY LIKE '" & Trim(RS1("INV_IDY")) & "' AND TXN_TYP LIKE '" & Trim(RS1("TXN_TYP")) & "'", con, adOpenKeyset, adLockPessimistic
    End If
    If rs2.RecordCount > 0 Then
        rs("cst_idy") = Trim(rs2("cst_idy"))
        rs("TXN_RMK") = Trim(rs2("TXN_RMK") & " ")
    End If
    rs.Update
    RS1.MoveNext
Next
'OPS UPDATION
'CON.Execute "delete from OPS"
'If rs.State = 1 Then rs.Close
'If RS1.State = 1 Then RS1.Close
'rs.Open "select * from OPS", CON, adOpenKeyset, adlockpessimistic
'RS1.Open "select * from INV,PRD,PGR WHERE STK_UPD ='Y' AND PRD.PRD_IDY LIKE INV.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND TXN_TYP IN ('PUR','SRN','STI','DMR','OPS')  and inv_dat < ctod('" & frm_dat & "')" & Trim(qry), CON, adOpenKeyset, adlockpessimistic
'For A = 1 To RS1.RecordCount
'    If rs.State = 1 Then rs.Close
'    rs.Open "SELECT * FROM OPS WHERE PRD_IDY LIKE '" & Trim(RS1("PRD_IDY")) & "'", CON, adOpenKeyset, adlockpessimistic
'    If rs.RecordCount > 0 Then
'        R_QTY = Val(MTS(Val(rs("PRD_QTY") & ""), RS1("SUB_UNT"))) + Val(MTS(RS1("PRD_QTY"), RS1("SUB_UNT")))
'        rs("PRD_QTY") = STM(Val(R_QTY), RS1("SUB_UNT"))
'        rs.Update
'    Else
'        rs.AddNew
'        rs("prd_idy") = Trim(RS1("prd_idy"))
'        rs("PRD_qty") = Trim(RS1("prd_qty"))
'        rs.Update
'    End If
'    RS1.MoveNext
'Next

    If rs.State = 1 Then rs.Close
    con.Execute "DELETE FROM CLD"
    rs.Open "select * from cld", con, adOpenKeyset, adLockPessimistic
    If RS1.State = 1 Then RS1.Close
    'RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat < ctod('" & FRM_DAT & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adlockpessimistic
    RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat < ctod('" & FRM_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic ' Modified by Radhika & Ramanesh
    For A = 1 To RS1.RecordCount
        If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!PRD_IDY = Trim(RS1!BAT_IDY)
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!sal_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!PRD_IDY = Trim(RS1!BAT_IDY)
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!sal_stk = Val(RS1(0)) * -1
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "OPS" Then
            con.Execute "UPDATE CLD SET OPN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!PRD_IDY = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = 0
                rs!opn_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "PUR" Then
        con.Execute "UPDATE CLD SET PUR_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!PRD_IDY = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!PUR_STK = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "PRT" Then
        con.Execute "UPDATE CLD SET PRT_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!PRD_IDY = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!PRT_STK = Val(RS1(0))
                rs!PUR_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "STI" Then
        con.Execute "UPDATE CLD SET STI_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!PRD_IDY = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs!STI_STK = Val(RS1(0))
                rs!PRT_STK = 0
                rs!PUR_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "SRN" Then
        con.Execute "UPDATE CLD SET SRN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!PRD_IDY = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!PUR_STK = 0
                rs!srn_stk = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs(7) = 0
                rs!cls_stk = 0
                rs.Update
            End If
        ElseIf RS1!TXN_TYP = "STO" Then
        con.Execute "UPDATE CLD SET STO_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!PRD_IDY = Trim(RS1!BAT_IDY)
                rs!sal_stk = 0
                rs!opn_stk = 0
                rs!srn_stk = 0
                rs(7) = Val(RS1(0))
                rs!PRT_STK = 0
                rs!STI_STK = 0
                rs!PUR_STK = 0
                rs!cls_stk = 0
                rs.Update
            End If
        End If
        RS1.MoveNext
    Next
    rs.Close
    con.Execute "UPDATE CLD SET CLS_STK=(PUR_STK+OPN_STK+SRN_STK+STI_STK)-(SAL_STK+PRT_STK+STO_STK)"
    con.Execute "UPDATE CLD SET OPN_STK=CLS_STK"
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
con.Execute "DELETE FROM OPS"
RS1.Open "select * from CLD,PRD,BRD WHERE CLD.PRD_IDY LIKE BRD.BAT_IDY AND PRD.PRD_IDY LIKE BRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
For A = 1 To RS1.RecordCount
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM OPS WHERE OPS.PRD_IDY LIKE '" & Trim(RS1("PRD_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        R_QTY = Val(RS1("CLS_STK") & "") + Val(MTS(rs("PRD_QTY"), RS1("SUB_UNT")))
        rs("PRD_QTY") = STM(Val(R_QTY), RS1("SUB_UNT"))
        rs.Update
    Else
        rs.AddNew
        rs("PRD_IDY") = Trim(RS1("PRD_IDY"))
        rs("PRD_QTY") = STM(Val(RS1("CLS_STK") & ""), Val(RS1!SUB_UNT))
        rs.Update
    End If
    RS1.MoveNext
Next
M_FRM_PPQ.MousePointer = vbDefault
If Trim(T) = "Stock Ledger Test" Then
    OpenReport CURDIR & "\Report.RP1", "Stock Ledger Consolidated", 0, Trim(SAL_QRY), 6
Else
    OpenReport CURDIR & "\Report.RP1", "Stock Ledger", 0, Trim(SAL_QRY), 6
End If
End Sub
Public Sub FAupdateCLG(FRM_DAT1 As Date, TO_DATE1 As Date, grpid As String)
FRM_DAT = FRM_DAT1
TO_DATE = TO_DATE1
M_FRM_APD.Tag = "FA_CLG"
M_FRM_APD.Show 1
End Sub
Public Sub FAupdate(FRM_DAT1 As Date, TO_DATE1 As Date, grpid As String, tp As String)
GROUPID = grpid
ACCOUNTTYP = tp
FRM_DAT = FRM_DAT1
TO_DATE = TO_DATE1
M_FRM_APD.Tag = "FA"
M_FRM_APD.Show 1
End Sub

Public Sub O_UPD_STK()
If MsgBox("Update Stocks..", vbExclamation + vbYesNo, "DAS Version 4.0") = vbYes Then
If CHECKTB("inv.dbf,prd.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

con.Execute "UPDATE PRD SET CUR_LEV = 0"
Dim RSI As New ADODB.Recordset
If RSI.State = 1 Then RSI.Close
RSI.Open "SELECT INV.TXN_TYP,INV.PRD_QTY,INV.PRD_IDY,PRD.SUB_UNT FROM INV,PRD WHERE INV.TXN_TYP IN ('OPS','PUR','SRN','STI','RSR') AND PRD.PRD_IDY LIKE INV.PRD_IDY", con, adOpenKeyset, adLockPessimistic
Do While Not RSI.EOF
    If Val(RSI(3)) <> 1 Then
        con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= " & ADDQTY(RSI(1), RSI(2))
    Else
        con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= CUR_LEV +" & RSI(1)
    End If
    RSI.MoveNext
Loop

If RSI.State = 1 Then RSI.Close
RSI.Open "SELECT INV.TXN_TYP,INV.PRD_QTY,INV.PRD_IDY,PRD.SUB_UNT FROM INV,PRD WHERE INV.TXN_TYP LIKE 'LSL' AND PRD.PRD_IDY LIKE INV.PRD_IDY", con, adOpenKeyset, adLockPessimistic
Do While Not RSI.EOF
    If Val(RSI(3)) <> 1 Then
        con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= " & CONQTY(RSI(1), RSI(2))
    Else
        con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= CUR_LEV- " & RSI(1)
    End If
    RSI.MoveNext
Loop

If RSI.State = 1 Then RSI.Close
RSI.Open "SELECT INV.TXN_TYP,INV.PRD_QTY,INV.PRD_IDY,PRD.SUB_UNT FROM INV,PRD WHERE INV.TXN_TYP IN ('SAL','RSI','STO','PRT') AND PRD.PRD_IDY LIKE INV.PRD_IDY", con, adOpenKeyset, adLockPessimistic
Do While Not RSI.EOF
    If Val(RSI(3)) <> 1 Then
        con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= " & CONQTY(RSI(1), RSI(2))
    Else
        con.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(RSI(2)) & "' SET CUR_LEV= CUR_LEV -" & RSI(1)
    End If
    RSI.MoveNext
Loop

End If
End Sub

Public Function MTS(M_QTY As Double, S_QTY As Integer) As Long
On Error GoTo LBL
If M_QTY < 0 Then
    M_CHK = True
Else
    M_CHK = False
End If
M_QTY = Abs(M_QTY)
If Val(S_QTY) <> 1 Then
    A = Int(M_QTY) * Val(S_QTY)
    b = (Val(M_QTY) - Int(M_QTY))
    b = b * 1000
    C = Val(A) + Val(b)
Else
    C = M_QTY
End If
If M_CHK = True Then C = C * -1
MTS = C
Exit Function
LBL:
MsgBox err.Description
End Function

Public Function STM(M_QTY As Double, S_QTY As Integer) As Double
If M_QTY < 0 Then
    M_CHK = True
Else
    M_CHK = False
End If
M_QTY = Abs(M_QTY)
If S_QTY <> 1 Then
    If S_QTY = 0 Then
        S_QTY = 1
        M_MOD = Val(M_QTY) Mod S_QTY
        A = (M_QTY - M_MOD)
        b = Val(A) / Val(S_QTY)
        C = Val(b) + (M_MOD / 1000)
    Else
        M_MOD = Val(M_QTY) Mod S_QTY
        A = (M_QTY - M_MOD)
        b = Val(A) / Val(S_QTY)
        C = Val(b) + (M_MOD / 1000)
    End If
Else
    C = M_QTY
End If
If M_CHK = True Then C = C * -1
STM = C
End Function

Public Sub createcompany(Destdir As String)
Dim FS As Object, f As Object, F1 As Object, FN As Object
Dim cn As New ADODB.Connection

If Dir(App.Path & "\bkp", vbDirectory) = "" Then
    MkDir App.Path & "\bkp"
    Destdir = App.Path & "\bkp\"
    Set FS = CreateObject("Scripting.FileSystemObject")
       
    SPATH = App.Path & "\???.dbf"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    SPATH = App.Path & "\eff.*"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    SPATH = App.Path & "\*.cdx"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    SPATH = App.Path & "\*.RP1"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    SPATH = App.Path & "\*.fpt"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    SPATH = App.Path & "\*.RRW"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    
    CopyFile App.Path & "\IND.EXE", Destdir & "IND.EXE", 0
    CopyFile App.Path & "\CND.EXE", Destdir & "CND.EXE", 0
    CopyFile App.Path & "\RRUNTIME.EXE", Destdir & "RRUNTIME.EXE", 0
    CopyFile App.Path & "\RRUN.EXE", Destdir & "RRUN.EXE", 0
    CopyFile App.Path & "\RR.CNF", Destdir & "RR.CNF", 0
    CopyFile App.Path & "\PRI.EXE", Destdir & "PRI.EXE", 0
    CopyFile App.Path & "\PRB.EXE", Destdir & "PRB.EXE", 0
    CopyFile App.Path & "\RRWRUN.EXE", Destdir & "RRWRUN.EXE", 0
    
    If cn.State = 1 Then cn.Close
    Set cn = Nothing
    cn.Open "Driver={Microsoft Visual Foxpro Driver};UID=;PWD=;SourceDB=" & Destdir & "Eff.dbc;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    
    Set f = FS.GetFolder(Destdir)
    Set FN = f.Files
    For Each F1 In FN
        tp = F1.Type
        If Right(Trim(Trim(F1.name)), 4) = ".dbf" Or Right(Trim(Trim(F1.name)), 4) = ".DBF" Then
            If Trim(F1.name) = "dir.dbf" Or Trim(F1.name) = "pre.dbf" Or Trim(F1.name) = "idy.dbf" Or Trim(F1.name) = "cat.dbf" Or Trim(F1.name) = "loc.dbf" Or Trim(F1.name) = "grp.dbf" Or Trim(F1.name) = "typ.dbf" Or Trim(F1.name) = "fun.dbf" Or Trim(F1.name) = "rri.dbf" Or Trim(F1.name) = "foxuser.dbf" Or Trim(F1.name) = "Dkcontrl.dbf" Or Trim(F1.name) = "Optional.dbf" Or Trim(F1.name) = "Required.dbf" Or Trim(F1.name) = "str_doc.dbf" Or Trim(F1.name) = "pgr.dbf" Or Trim(F1.name) = "Reports.dbf" Or Trim(F1.name) = "yer.dbf" Or Trim(F1.name) = "pre.dbf" Then
            ElseIf Trim(F1.name) = "usr.dbf" Then
                cn.Execute "delete from " & Trim(Trim(F1.name)) & " where usr_nme<>'ADMINISTRATOR'"
            Else
                cn.Execute "delete from " & Trim(Trim(F1.name))
            End If
        End If
    Next
    
    If cn.State = 1 Then cn.Close
    Set cn = Nothing
    cn.Open "Driver={Microsoft Visual Foxpro Driver};UID=;PWD=;SourceDB=" & Destdir & "Eff.dbc;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    
    cn.Execute "update idy set idy_sal='0'"
    cn.Execute "update idy set idy_act='0'"
    cn.Execute "update idy set idy_prd='0'"
    cn.Execute "update idy set idy_pgr='0'"
    cn.Execute "update idy set idy_pur='0'"
    cn.Execute "update idy set idy_srn='0'"
    cn.Execute "update idy set idy_prt='0'"
    cn.Execute "update idy set idy_ops='0'"
    cn.Execute "update idy set idy_rct='0'"
    cn.Execute "update idy set idy_pay='0'"
    cn.Execute "update idy set idy_jrn='0'"
    cn.Execute "update idy set idy_opn='0'"
    cn.Execute "update idy set idy_rsi='0'"
    cn.Execute "update idy set idy_mar='0'"
    cn.Execute "update idy set idy_smn='0'"
    cn.Execute "update idy set idy_chn='0'"
    cn.Execute "update idy set idy_ord='0'"
    cn.Execute "update idy set idy_otc='0'"
    cn.Execute "update idy set idy_irt='0'"
    cn.Execute "update idy set idy_CRN='0'"
    cn.Execute "update idy set idy_DBN='0'"
    cn.Execute "update idy set idy_STI='0'"
    cn.Execute "update idy set idy_STO='0'"
    cn.Execute "update idy set idy_LDS='0'"
    cn.Execute "update idy set idy_SCH='0'"
    cn.Execute "update idy set idy_BAT='0'"
    cn.Execute "update idy set idy_TXN='0'"
    cn.Execute "update idy set idy_SHL='0'"
    cn.Execute "update idy set idy_SHT='0'"
    cn.Execute "update idy set idy_adj='0'"
    cn.Execute "update fun set fun_two=''"
    cn.Execute "update fun set fun_thr=''"
    cn.Execute "update fun set fun_for=''"
    cn.Execute "update fun set fun_fiv=''"
    cn.Execute "update fun set fun_six=''"
    cn.Execute "update fun set fun_sev=''"
    cn.Execute "update fun set fun_eig=''"
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

Public Function PERCENT(amt As String, PRC As String) As String
PERCENT = Round(Val(amt) * (Val(PRC) / 100), 2)
End Function
Public Function PERCADD(amt As String, PRC As String) As String
PCGADD = (Val(amt) * (Val(PRC) / 100)) + amt
End Function
Public Function PERCLESS(amt As String, PRC As String) As String
PCGADD = amt - (Val(amt) * (Val(PRC) / 100))
End Function

Public Function GetValue(fldnames As String, tblName As String) As String
'If CHECKTB(tblname & ".dbf") = 1 Then
'    MsgBox UCase(tblname) & " Missing "
'    Exit Function
'End If
Dim rs As New ADODB.Recordset
rs.Open "select " & fldnames & " from " & tblName, con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    GetValue = Trim(rs(0) & "")
End If
End Function
Public Sub TRAIL_BALANCECLG(TO_DAT1 As Date)
TO_DAT = TO_DAT1
Unload M_FRM_ASD
M_FRM_APD.Tag = "TB_CLG"
M_FRM_APD.Show 1
End Sub
Public Sub inltopnl()
On Error GoTo ERRH1
If CHECKTB("inl.dbf,pnl.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
rs.Open "select * from inl", CN1, adOpenKeyset, adLockPessimistic
RS1.Open "select * from pnl", CN1, adOpenKeyset, adLockPessimistic
CN1.Execute "delete from pnl"
While Not rs.EOF
RS1.AddNew
For i = 0 To rs.Fields.count - 1
FLD = rs.Fields(i).name
RS1(FLD) = rs(FLD)
Next
RS1.Update
rs.MoveNext
RS1.CancelUpdate
Wend
'MsgBox "completed"
If rs.RecordCount > 0 Then
rs.MoveFirst
While Not rs.EOF
For i = 4 To rs.Fields.count - 1
If rs.Fields(i).Type = adInteger Or rs.Fields(i).Type = adNumeric Or rs.Fields(i).Type = adBigInt Then
CN1.Execute "update pnl set " & rs.Fields(i).name & "=" & rs(i) & "where inv_idy like '" & rs("inv_idy") & "'"
ElseIf rs.Fields(i).Type = adDate Or rs.Fields(i).Type = adDBDate Or rs.Fields(i).Type = adDBTimeStamp Then
CN1.Execute "update pnl set " & rs.Fields(i).name & "=ctod('" & rs(i) & "') where inv_idy like '" & rs("inv_idy") & "'"
Else
CN1.Execute "update pnl set " & rs.Fields(i).name & "='" & rs(i) & "' where inv_idy like '" & rs("inv_idy") & "'"
End If
Next
rs.MoveNext
'MsgBox "comp"
Wend
End If
'MsgBox "Full"
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
Exit Sub
ERRH1:
If err.Number = 3265 Then
    Resume Next
ElseIf err.Number = -2147217887 Then
If RS1(FLD).Type = adChar Or RS1(FLD).Type = adVarChar Or RS1(FLD).Type = adBSTR Or RS1(FLD).Type = adLongVarChar Then
RS1(FLD) = ""
Resume Next
ElseIf RS1(FLD).Type = adNumeric Or RS1(FLD).Type = adBigInt Or RS1(FLD).Type = adInteger Or RS1(FLD).Type = adSingle Or RS1(FLD).Type = adDouble Then
RS1(FLD) = 0
Resume Next
ElseIf RS1(FLD).Type = adDate Or RS1(FLD).Type = adDBDate Or RS1(FLD).Type = adDBTimeStamp Then
RS1(FLD) = Date
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
ElseIf RS1(fnm).Type = adNumeric Or RS1(fnm).Type = adBigInt Or RS1(fnm).Type = adInteger Or RS1(fnm).Type = adSingle Or RS1(FLD).Type = adDouble Then
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
MsgBox err.Description & " " & err.Number & "  " & FLD
End If
End Sub
Public Sub invtopnv()
On Error GoTo errh2
If CHECKTB("inv.dbf,pnv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
rs.Open "select * from inv", CN1, adOpenKeyset, adLockPessimistic
RS1.Open "select * from pnv", CN1, adOpenKeyset, adLockPessimistic
CN1.Execute "delete from pnv"
While Not rs.EOF
RS1.AddNew
For i = 0 To rs.Fields.count - 1
FLD = rs.Fields(i).name
RS1(FLD) = rs(FLD)
Next
RS1.Update
rs.MoveNext
RS1.CancelUpdate
Wend
'MsgBox "completed"
If rs.RecordCount > 0 Then
rs.MoveFirst
While Not rs.EOF
For i = 3 To rs.Fields.count - 1
If rs.Fields(i).Type = adInteger Or rs.Fields(i).Type = adNumeric Or rs.Fields(i).Type = adBigInt Then
CN1.Execute "update pnv set " & rs.Fields(i).name & "=" & rs(i) & "where inv_idy like '" & rs("inv_idy") & "'"
ElseIf rs.Fields(i).Type = adDate Or rs.Fields(i).Type = adDBDate Or rs.Fields(i).Type = adDBTimeStamp Then
CN1.Execute "update pnv set " & rs.Fields(i).name & "=ctod('" & rs(i) & "') where inv_idy like '" & rs("inv_idy") & "'"
Else
CN1.Execute "update pnv set " & rs.Fields(i).name & "='" & rs(i) & "' where inv_idy like '" & rs("inv_idy") & "'"
End If
Next
rs.MoveNext
'MsgBox "comp"
Wend
End If
'MsgBox "Full"
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
Exit Sub
errh2:
If err.Number = 3265 Then
    Resume Next
ElseIf err.Number = -2147217887 Then
If RS1(FLD).Type = adChar Or RS1(FLD).Type = adVarChar Or RS1(FLD).Type = adBSTR Or RS1(FLD).Type = adLongVarChar Then
RS1(FLD) = ""
Resume Next
ElseIf RS1(FLD).Type = adNumeric Or RS1(FLD).Type = adBigInt Or RS1(FLD).Type = adInteger Or RS1(FLD).Type = adSingle Or RS1(FLD).Type = adDouble Then
RS1(FLD) = 0
Resume Next
ElseIf RS1(FLD).Type = adDate Or RS1(FLD).Type = adDBDate Or RS1(FLD).Type = adDBTimeStamp Then
RS1(FLD) = Date
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
ElseIf RS1(fnm).Type = adNumeric Or RS1(fnm).Type = adBigInt Or RS1(fnm).Type = adInteger Or RS1(fnm).Type = adSingle Or RS1(FLD).Type = adDouble Then
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
MsgBox err.Description & " " & err.Number & "  " & FLD
End If
End Sub


Public Sub prdtoinv()
If CHECKTB("prd.dbf,inv.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing "
    Exit Sub
End If

If rs2.State = 1 Then rs2.Close
    If rs3.State = 1 Then rs3.Close
    rs3.Open "SELECT * FROM PRD", CN1, adOpenDynamic, adLockPessimistic
    While Not rs3.EOF
        CN1.Execute "update inv set sub_unt = " & Val(rs3!SUB_UNT) & " where prd_idy like '" & Trim(rs3!PRD_IDY) & "'", d
        CN1.Execute "update inv set prd_uom = " & Val(rs3!con_fac) & " where prd_idy like '" & Trim(rs3!PRD_IDY) & "'", d
        rs3.MoveNext
    Wend
    CN1.Execute "update inv set stk_typ='Salable'"
    CN1.Execute "update inl set txn_typ = 'SAL' WHERE TXN_TYP LIKE 'LSL'"
    CN1.Execute "update inv set txn_typ = 'SAL' WHERE TXN_TYP LIKE 'LSL'"
    'CN1.Execute "update inv set stk_typ='Salable'"
    CN1.Execute "update inv set stk_typ = 'Free' WHERE txn_typ IN('FRC','FRS')"
    CN1.Execute "update inv set prd_unt=int(prd_qty)"
    CN1.Execute "update inv set prd_pcs=(prd_qty-int(prd_qty))*1000"
    CN1.Execute "update inv set tot_pcs=(prd_qty * sub_unt) + prd_pcs"
    CN1.Execute "update inv set prd_uom=(prd_uom/sub_unt) * tot_pcs"
    If rs2.State = 1 Then rs2.Close
    If rs3.State = 1 Then rs3.Close
End Sub

Public Function CHKSUBUNT(STK As Double, SUBUNT As Integer) As Boolean
S_UNT = (Val(STK) - Int(STK)) * 1000
If S_UNT < SUBUNT Then
    CHKSUBUNT = True
Else
    MsgBox "Product Subunits are " & SUBUNT, vbExclamation
    CHKSUBUNT = False
End If
End Function
Public Function CHECKTB(sFile)
tbN = Split(sFile, ",", , vbTextCompare)
For b = 0 To UBound(tbN)
    strPath = CURDIR & "\" & Trim(tbN(b))
    If Dir(strPath) = "" Then
        CHECKTB = 1
        TNAME = tbN(b)
        Exit Function
    End If
Next
End Function
Public Sub SLT()
SendKeys "{HOME}+{END}"
End Sub

Public Function GETROUTE(CustomerID As String, SalesmanID, BillDate As Date) As String
Dim rs As New ADODB.Recordset
Select Case Format(BillDate, "DDDD")
    Case "Monday"
        d = 1
    Case "Tuesday"
        d = 2
    Case "Wednesday"
        d = 3
    Case "Thursday"
        d = 4
    Case "Friday"
        d = 5
    Case "Saturday"
        d = 6
    Case "Sunday"
        d = 7
End Select
rs.Open "select * from rdy where act_idy like '" & Trim(CustomerID) & "' and smn_idy like '" & Trim(SalesmanID) & "' and day_idy =" & d, con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then GETROUTE = Trim(rs!rou_idy & "")
End Function
Public Function GetSalesman(CustomerID As String, BillDate As Date) As String
Dim rs As New ADODB.Recordset
Select Case Format(BillDate, "DDDD")
    Case "Monday"
        d = 1
    Case "Tuesday"
        d = 2
    Case "Wednesday"
        d = 3
    Case "Thursday"
        d = 4
    Case "Friday"
        d = 5
    Case "Saturday"
        d = 6
    Case "Sunday"
        d = 7
End Select
rs.Open "select * from rdy,smn where smn.smn_idy like rdy.smn_idy and rdy.act_idy like '" & Trim(CustomerID) & "' and day_idy =" & d, con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then GetSalesman = Trim(rs!SMN_IDY & "") & "|" & Trim(rs!SMN_NME & "")
End Function

'Public Sub AddShells(Invoice As String, Txntype As String, PartyID As String, UPD As String)
'Dim RSINV As New ADODB.Recordset
'Dim RSSNV As New ADODB.Recordset
'CON.Execute "DELETE FROM SNV WHERE INV_IDY LIKE '" & Trim(Invoice) & "' AND TXN_TYP LIKE '" & Trim(Txntype) & "'"
'If Txntype = "SAL" Then
'    RSINV.Open "SELECT sum(tot_pcs),inv.prd_idy,inv_idy,inv_dat,shl_idy,sub_unt FROM INV,SHL WHERE INV.PRD_IDY LIKE SHL.PRD_IDY AND INV_IDY LIKE '" & Trim(Invoice) & "' AND TXN_TYP LIKE '" & Trim(Txntype) & "' AND SHL_TYP ='Shell' AND STK_TYP IN('Salable','Free') group by inv.prd_idy,inv.stk_typ", CON, adOpenKeyset, adlockpessimistic
'Else
'    RSINV.Open "SELECT sum(tot_pcs),inv.prd_idy,inv_idy,inv_dat,shl_idy,sub_unt FROM INV,SHL WHERE INV.PRD_IDY LIKE SHL.PRD_IDY AND INV_IDY LIKE '" & Trim(Invoice) & "' AND TXN_TYP LIKE '" & Trim(Txntype) & "' AND SHL_TYP ='Shell' group by inv.prd_idy", CON, adOpenKeyset, adlockpessimistic
'End If
'RSSNV.Open "SELECT * FROM SNV WHERE INV_IDY LIKE ''", CON, adOpenKeyset, adlockpessimistic
'For A = 1 To RSINV.RecordCount
'    If Txntype = "SAL" Then If A = 1 Then prdid = RSINV!PRD_IDY
'    RSSNV.AddNew
'    RSSNV!inv_idy = Trim(RSINV!inv_idy)
'    RSSNV!INV_DAT = Trim(RSINV!INV_DAT)
'    RSSNV!PRD_IDY = Trim(RSINV!PRD_IDY)
'    RSSNV!TXN_TYP = Trim(Txntype)
'    RSSNV!shl_idy = Trim(RSINV!shl_idy)
'    RSSNV!pck_idy = ""
'    RSSNV!i_shl_nos = Int(Val(STM(Val(RSINV(0) & ""), Val(RSINV(5)))))
'    RSSNV!i_pck_nos = Val(RSINV(0))
'    RSSNV!R_shl_nos = Val(RSSNV!i_shl_nos)
'    RSSNV!R_pck_nos = Val(RSSNV!i_pck_nos & "")
'    RSSNV!D_shl_nos = 0
'    RSSNV!D_pck_nos = 0
'    RSSNV!dp_shl_nos = 0
'    RSSNV!dp_pck_nos = 0
'    RSSNV!dep_amt = 0
'    RSSNV!brk_shl = 0
'    RSSNV!brk_pck = 0
'    RSSNV!txn_dat = Trim(RSINV!INV_DAT)
'    RSSNV!TXN_IDY = ""
'    RSSNV!CST_IDY = Trim(PartyID)
'    RSSNV!stk_upd = UPD
'    RSSNV.Update
'    RSINV.MoveNext
'Next
'If Txntype = "SAL" Then
'    If RSINV.State = 1 Then RSINV.Close
'    RSINV.Open "SELECT sum(tot_pcs),inv.prd_idy,inv_idy,inv_dat,shl_idy,sub_unt FROM INV,SHL WHERE INV.PRD_IDY LIKE SHL.PRD_IDY AND INV_IDY LIKE '" & Trim(Invoice) & "' AND TXN_TYP LIKE '" & Trim(Txntype) & "' AND SHL_TYP ='Shell' AND STK_TYP IN ('Breakage') group by inv.prd_idy", CON, adOpenKeyset, adlockpessimistic
'    For A = 1 To RSINV.RecordCount
'        CON.Execute "update snv set r_shl_nos=r_shl_nos-" & Int(STM(RSINV(0), RSINV(5))) & ",r_pck_nos=r_pck_nos-" & Val(RSINV(0)) & ",brk_shl=" & Int(STM(RSINV(0), RSINV(5))) & ",brk_pck=" & Val(RSINV(0)) & "where prd_idy like '" & Trim(RSINV(1)) & "' and inv_idy like '" & Trim(Invoice) & "' and txn_typ like '" & Trim(Txntype) & "'", d
'        'CON.Execute "update snv set r_shl_nos=i_shl_nos,r_pck_nos=i_pck_nos where prd_idy like '" & Trim(RSINV(1)) & "' and inv_idy like '" & Trim(Invoice) & "' and txn_typ like '" & Trim(Txntype) & "'"
'        If d = 0 Then
'            CON.Execute "update snv set r_shl_nos=r_shl_nos-" & Int(STM(RSINV(0), RSINV(5))) & ",r_pck_nos=r_pck_nos-" & Val(RSINV(0)) & " where prd_idy like '" & Trim(prdid) & "' and inv_idy like '" & Trim(Invoice) & "' and txn_typ like '" & Trim(Txntype) & "'", d
'            RSSNV.AddNew
'            RSSNV!inv_idy = Trim(RSINV(2))
'            RSSNV!INV_DAT = Trim(RSINV(3))
'            RSSNV!shl_idy = Trim(RSINV(4))
'            RSSNV!TXN_TYP = Trim(Txntype)
'            RSSNV!pck_idy = ""
'            'rssnv!i_shl_nos = STM(RSINV(0), RSINV(5)) * -1
'            'rssnv!i_pck_nos = Val(RSINV(0)) * -1
'            RSSNV!i_shl_nos = 0
'            RSSNV!i_pck_nos = 0
'            RSSNV!brk_shl = STM(RSINV(0), RSINV(5))
'            RSSNV!brk_pck = Val(RSINV(0))
'            RSSNV!CST_IDY = PartyID
'            RSSNV!PRD_IDY = Trim(RSINV!PRD_IDY)
'            RSSNV!txn_dat = Trim(RSINV!INV_DAT)
'            RSSNV!stk_upd = UPD
'            RSSNV.Update
'        End If
'        RSINV.MoveNext
'    Next
'End If
'End Sub
Public Sub AddShells(Invoice As String, Txntype As String, PartyID As String, UPD As String, M_SMN_IDY As String, M_ROU_IDY As String)
Dim RSINV As New ADODB.Recordset
Dim RSSNV As New ADODB.Recordset
con.Execute "DELETE FROM SNV WHERE INV_IDY LIKE '" & Trim(Invoice) & "' AND TXN_TYP LIKE '" & Trim(Txntype) & "'"
If Txntype = "SAL" Then
    RSINV.Open "SELECT sum(tot_pcs),inv.prd_idy,inv.inv_idy,inv.inv_dat,shl.shl_idy,prd.sub_unt,prd.prd_pck,INV.ORD_IDY FROM INV,SHL,PRD WHERE PRD.PRD_IDY LIKE INV.PRD_IDY AND INV.PRD_IDY LIKE SHL.PRD_IDY AND INV_IDY LIKE '" & Trim(Invoice) & "' AND inv.TXN_TYP LIKE '" & Trim(Txntype) & "' AND SHL_TYP ='Shell' AND INV.STK_TYP IN('Salable','Free') group by inv.prd_idy,inv.stk_typ", con, adOpenKeyset, adLockPessimistic
Else
    RSINV.Open "SELECT sum(tot_pcs),inv.prd_idy,INV.inv_idy,INV.inv_dat,SHL.shl_idy,PRD.sub_unt,prd.PRD_Pck,INV.ORD_IDY FROM INV,SHL,PRD WHERE PRD.PRD_IDY LIKE INV.PRD_IDY AND INV.PRD_IDY LIKE SHL.PRD_IDY AND INV_IDY LIKE '" & Trim(Invoice) & "' AND INV.TXN_TYP LIKE '" & Trim(Txntype) & "' AND SHL_TYP ='Shell' AND INV.STK_TYP NOT IN('Breakage') group by inv.prd_idy", con, adOpenKeyset, adLockPessimistic
End If
RSSNV.Open "SELECT * FROM SNV WHERE INV_IDY LIKE ''", con, adOpenKeyset, adLockPessimistic
For A = 1 To RSINV.RecordCount
    If Txntype = "SAL" Then If A = 1 Then PRDID = RSINV!PRD_IDY         ' TO DISCUSS WITH REDDY SIR
    RSSNV.AddNew
    RSSNV!SMN_IDY = Trim(M_SMN_IDY)
    RSSNV!rou_idy = Trim(M_ROU_IDY)
    RSSNV!INV_IDY = Trim(RSINV!INV_IDY)
    RSSNV!INV_DAT = Trim(RSINV!INV_DAT)
    RSSNV!PRD_IDY = Trim(RSINV!PRD_IDY)
    RSSNV!TXN_TYP = Trim(Txntype)
    RSSNV!shl_idy = Trim(RSINV!shl_idy)
    RSSNV!pck_idy = ""
    If Txntype = "PUR" Or Txntype = "STI" Or Txntype = "STO" Then
        RSSNV!I_shl_nos = Int(Val(STM(Val(RSINV(0) & ""), Val(RSINV(5)))))
        RSSNV!I_pck_nos = Val(RSINV(0))
        RSSNV!R_shl_nos = 0#
        RSSNV!R_pck_nos = 0#
        RSSNV!d_shl_nos = (RSSNV!I_shl_nos - RSSNV!R_shl_nos)
        RSSNV!d_pck_nos = (RSSNV!I_pck_nos - RSSNV!R_pck_nos)
    Else
        RSSNV!I_shl_nos = Int(Val(STM(Val(RSINV(0) & ""), Val(RSINV(5)))))
        RSSNV!I_pck_nos = Val(RSINV(0))
        RSSNV!R_shl_nos = Val(RSSNV!I_shl_nos)
        RSSNV!R_pck_nos = Val(RSSNV!I_pck_nos & "")
        RSSNV!d_shl_nos = 0
        RSSNV!d_pck_nos = 0
    End If
    RSSNV!dp_shl_nos = 0
    RSSNV!dp_pck_nos = 0
    RSSNV!dep_amt = 0
    RSSNV!BRK_SHL = 0
    RSSNV!BRK_PCK = 0
    RSSNV!txn_dat = Trim(RSINV!INV_DAT)
    RSSNV!TXN_IDY = Trim(RSINV!ORD_IDY & "")
    RSSNV!CST_IDY = Trim(PartyID)
    RSSNV!PRD_PCK = Trim(RSINV(6) & "")
    RSSNV!prd_upc = Val(RSINV!SUB_UNT)
    RSSNV!stk_upd = UPD
    RSSNV.Update
    RSINV.MoveNext
Next

If Txntype = "SAL" Then
    If RSINV.State = 1 Then RSINV.Close
    RSINV.Open "SELECT sum(tot_pcs),inv.prd_idy,inv.inv_idy,inv.inv_dat,shl.shl_idy,prd.sub_unt,prd.prd_pck,INV.ORD_IDY FROM INV,SHL,PRD WHERE PRD.PRD_IDY LIKE INV.PRD_IDY AND INV.PRD_IDY LIKE SHL.PRD_IDY AND INV_IDY LIKE '" & Trim(Invoice) & "' AND inv.TXN_TYP LIKE '" & Trim(Txntype) & "' AND SHL_TYP ='Shell' AND INV.STK_TYP IN('Damage','TakeBack','Breakage','Leakage') group by inv.prd_idy", con, adOpenKeyset, adLockPessimistic
    For A = 1 To RSINV.RecordCount
        con.Execute "update snv set i_shl_nos = i_shl_nos - " & Int(Val(STM(Val(RSINV(0) & ""), Val(RSINV(5))))) & " where inv_idy = '" & Trim(RSINV!INV_IDY) & "' and prd_idy  ='" & RSINV!PRD_IDY & "'", d
        con.Execute "update snv set i_pck_nos = i_pck_nos - " & Val(RSINV(0) & "") & " where inv_idy = '" & Trim(RSINV!INV_IDY) & "' and prd_idy  ='" & RSINV!PRD_IDY & "'", e
        con.Execute "update snv set r_shl_nos = r_shl_nos - " & Int(Val(STM(Val(RSINV(0) & ""), Val(RSINV(5))))) & " where inv_idy = '" & Trim(RSINV!INV_IDY) & "' and prd_idy  ='" & RSINV!PRD_IDY & "'", f
        con.Execute "update snv set r_pck_nos = r_pck_nos - " & Val(RSINV(0) & "") & " where inv_idy = '" & Trim(RSINV!INV_IDY) & "' and prd_idy  ='" & RSINV!PRD_IDY & "'", g
        If d = 0 Or e = 0 Or f = 0 Or g = 0 Then
            RSSNV.AddNew
            RSSNV!SMN_IDY = Trim(M_SMN_IDY)
            RSSNV!rou_idy = Trim(M_ROU_IDY)
            RSSNV!INV_IDY = Trim(RSINV!INV_IDY)
            RSSNV!INV_DAT = Trim(RSINV!INV_DAT)
            RSSNV!PRD_IDY = Trim(RSINV!PRD_IDY)
            RSSNV!TXN_TYP = Trim(Txntype)
            RSSNV!shl_idy = Trim(RSINV!shl_idy)
            RSSNV!pck_idy = ""
            RSSNV!I_shl_nos = Int(Val(STM(Val(RSINV(0) & ""), Val(RSINV(5))))) * -1
            RSSNV!I_pck_nos = Val(RSINV(0)) * -1
            RSSNV!R_shl_nos = Val(RSSNV!I_shl_nos) * -1
            RSSNV!R_pck_nos = Val(RSSNV!I_pck_nos & "") * -1
            RSSNV!d_shl_nos = 0
            RSSNV!d_pck_nos = 0
            RSSNV!dp_shl_nos = 0
            RSSNV!dp_pck_nos = 0
            RSSNV!dep_amt = 0
            RSSNV!BRK_SHL = 0
            RSSNV!BRK_PCK = 0
            RSSNV!txn_dat = Trim(RSINV!INV_DAT)
            RSSNV!TXN_IDY = Trim(RSINV!ORD_IDY & "")
            RSSNV!CST_IDY = Trim(PartyID)
            RSSNV!PRD_PCK = Trim(RSINV(6) & "")
            RSSNV!prd_upc = Val(RSINV!SUB_UNT)
            RSSNV!stk_upd = UPD
            RSSNV.Update
        End If
        RSINV.MoveNext
    Next
End If
If RSINV.State = 1 Then RSINV.Close
RSINV.Open "SELECT sum(tot_pcs),inv.prd_idy,INV.inv_idy,INV.inv_dat,SHL.shl_idy,PRD.sub_unt,prd.prd_pck FROM INV,SHL,prd WHERE inv.prd_idy like prd.prd_idy and INV.PRD_IDY LIKE SHL.PRD_IDY AND INV_IDY LIKE '" & Trim(Invoice) & "' AND TXN_TYP LIKE '" & Trim(Txntype) & "' AND SHL_TYP ='Shell' AND INV.STK_TYP IN ('Breakage') group by inv.prd_idy", con, adOpenKeyset, adLockPessimistic
For A = 1 To RSINV.RecordCount
    con.Execute "update snv set r_shl_nos=r_shl_nos - " & Int(STM(RSINV(0), RSINV(5))) & ",r_pck_nos=r_pck_nos-" & Val(RSINV(0)) & ",brk_shl=" & Int(STM(RSINV(0), RSINV(5))) & ",brk_pck=" & Val(RSINV(0)) & "where prd_idy like '" & Trim(RSINV(1)) & "' and inv_idy like '" & Trim(Invoice) & "' and txn_typ like '" & Trim(Txntype) & "'", d
    'CON.Execute "update snv set r_shl_nos=i_shl_nos,r_pck_nos=i_pck_nos where prd_idy like '" & Trim(RSINV(1)) & "' and inv_idy like '" & Trim(Invoice) & "' and txn_typ like '" & Trim(Txntype) & "'"
    If d = 0 Then
        con.Execute "update snv set r_shl_nos=r_shl_nos-" & Int(STM(RSINV(0), RSINV(5))) & ",r_pck_nos=r_pck_nos-" & Val(RSINV(0)) & " where prd_idy like '" & Trim(PRDID) & "' and inv_idy like '" & Trim(Invoice) & "' and txn_typ like '" & Trim(Txntype) & "'", d
        RSSNV.AddNew
        RSSNV!INV_IDY = Trim(RSINV(2))
        RSSNV!INV_DAT = Trim(RSINV(3))
        RSSNV!shl_idy = Trim(RSINV(4))
        RSSNV!TXN_TYP = Trim(Txntype)
        RSSNV!pck_idy = ""
        'rssnv!i_shl_nos = STM(RSINV(0), RSINV(5)) * -1
        'rssnv!i_pck_nos = Val(RSINV(0)) * -1
        RSSNV!I_shl_nos = 0
        RSSNV!I_pck_nos = 0
        RSSNV!BRK_SHL = STM(RSINV(0), RSINV(5))
        RSSNV!BRK_PCK = Val(RSINV(0))
        RSSNV!CST_IDY = PartyID
        RSSNV!PRD_IDY = Trim(RSINV!PRD_IDY)
        RSSNV!txn_dat = Trim(RSINV!INV_DAT)
        RSSNV!PRD_PCK = Trim(RSINV!PRD_PCK)
        RSSNV!prd_upc = Val(RSINV!SUB_UNT)
        RSSNV!stk_upd = UPD
        RSSNV.Update
    End If
    RSINV.MoveNext
Next
End Sub

Public Sub SCl(M As MSFlexGrid, C As Integer, Optional r = 0)
M.Row = r
M.Col = C
M.CellForeColor = vbBlue
End Sub
Public Function CASEAMT(QT As String, SU As String, amt As String)
FD = InStr(1, Trim(QT), ".")
If FD <> 0 Then
    TC = Mid(Trim(QT), 1, FD - 1)
    SB = Mid(Trim(QT), FD + 1, Len(Trim(QT)))
    PU = Val(amt) / Val(SU)
    CASEAMT = (Val(TC) * Val(amt)) + (Val(SB) * Val(PU))
Else
    CASEAMT = Val(Trim(QT)) * Val(amt)
End If
End Function
Public Sub FillScheduledCalls()
Dim RST_RDY As New ADODB.Recordset
Dim RST_SMN As New ADODB.Recordset
If RST_SMN.State = 1 Then RST_SMN.Close
RST_SMN.Open "SELECT * FROM SMN", con, adOpenKeyset, adLockPessimistic
If RST_SMN.RecordCount > 0 Then RST_SMN.MoveFirst
While Not RST_SMN.EOF
    If RST_RDY.State = 1 Then RST_RDY.Close
'    RST_RDY.Open "SELECT COUNT(ACT_IDY), DAY_IDY FROM RDY,ACT WHERE SMN_IDY LIKE '" & RST_SMN("SMN_IDY") & "' GROUP BY SMN_IDY, DAY_IDY", CON, adOpenKeyset, adlockpessimistic
    RST_RDY.Open "SELECT DISTINCT COUNT(RDY.ACT_IDY), DAY_IDY FROM RDY,ACT WHERE RDY.SMN_IDY LIKE '" & RST_SMN("SMN_IDY") & "' AND RDY.ACT_IDY LIKE ACT.ACT_IDY AND ACT.DBT_TYP = 'Y' GROUP BY RDY.SMN_IDY, DAY_IDY", con, adOpenKeyset, adLockPessimistic
    If RST_RDY.RecordCount > 0 Then
        While Not RST_RDY.EOF
            If Val(RST_RDY("DAY_IDY")) = 1 Then
                RST_SMN("SMN_MON") = Val(RST_RDY(0))
            End If
            If Val(RST_RDY("DAY_IDY")) = 2 Then
                RST_SMN("SMN_TUE") = Val(RST_RDY(0))
            End If
            If Val(RST_RDY("DAY_IDY")) = 3 Then
                RST_SMN("SMN_WED") = Val(RST_RDY(0))
            End If
            If Val(RST_RDY("DAY_IDY")) = 4 Then
                RST_SMN("SMN_THU") = Val(RST_RDY(0))
            End If
            If Val(RST_RDY("DAY_IDY")) = 5 Then
                RST_SMN("SMN_FRI") = Val(RST_RDY(0))
            End If
            If Val(RST_RDY("DAY_IDY")) = 6 Then
                RST_SMN("SMN_SAT") = Val(RST_RDY(0))
            End If
            If Val(RST_RDY("DAY_IDY")) = 7 Then
                RST_SMN("SMN_SUN") = Val(RST_RDY(0))
            End If
            RST_SMN.Update
            RST_RDY.MoveNext
        Wend
    End If
    RST_SMN.MoveNext
Wend
End Sub
Public Sub AddScheduledCalls(FrmDate As Date, ToDate As Date, SMNorROUorACTorDAY As String)
Dim RSSMN As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim RSINV As New ADODB.Recordset
Dim RSRP1 As New ADODB.Recordset
Dim DT1 As Date
Dim DT2 As Date
Dim DAYCOUNT(6) As Integer
CON_APP = SMNorROUorACTorDAY
DT1 = FrmDate
DT2 = ToDate
con.Execute "DELETE FROM RP1"
Do While DT1 <= DT2
    If Format(DT1, "DDDD") = "Monday" Then
        DAYCOUNT(0) = Val(DAYCOUNT(0)) + 1
    ElseIf Format(DT1, "DDDD") = "Tuesday" Then
        DAYCOUNT(1) = Val(DAYCOUNT(1)) + 1
    ElseIf Format(DT1, "DDDD") = "Wednesday" Then
        DAYCOUNT(2) = Val(DAYCOUNT(2)) + 1
    ElseIf Format(DT1, "DDDD") = "Thursday" Then
        DAYCOUNT(3) = Val(DAYCOUNT(3)) + 1
    ElseIf Format(DT1, "DDDD") = "Friday" Then
        DAYCOUNT(4) = Val(DAYCOUNT(4)) + 1
    ElseIf Format(DT1, "DDDD") = "Saturday" Then
        DAYCOUNT(5) = Val(DAYCOUNT(5)) + 1
    ElseIf Format(DT1, "DDDD") = "Sunday" Then
        DAYCOUNT(6) = Val(DAYCOUNT(6)) + 1
    End If
    DT1 = DateAdd("D", 1, DT1)
Loop

If RSRP1.State = 1 Then RSRP1.Close
RSRP1.Open "SELECT * FROM RP1", con, adOpenKeyset, adLockPessimistic

If CON_APP = "SMN" Then
    If RSSMN.State = 1 Then RSSMN.Close
    RSSMN.Open "SELECT * FROM SMN", con, adOpenKeyset
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT COUNT(*),SMN_IDY,DAY_IDY FROM RDY GROUP BY DAY_IDY,SMN_IDY", con, adOpenKeyset, adLockPessimistic
    For A = 1 To RSSMN.RecordCount
        CTR = 0
        For dayid = 0 To 6
            rs.filter = "SMN_IDY LIKE '" & Trim(RSSMN!SMN_IDY) & "' AND DAY_IDY =" & dayid + 1
            If rs.RecordCount > 0 Then CTR = CTR + (Val(rs(0)) * DAYCOUNT(dayid))
        Next
        RSRP1.AddNew
        RSRP1(0) = ""
        RSRP1(1) = ""
        RSRP1(2) = ""
        RSRP1(3) = ""
        RSRP1(4) = ""
        RSRP1(5) = ""
        RSRP1(6) = ""
        RSRP1(7) = ""
        RSRP1(8) = ""
        RSRP1(9) = RSSMN!SMN_IDY
        RSRP1(10) = 0
        RSRP1(11) = CTR
        RSRP1(12) = 0
        RSRP1(13) = 0
        RSRP1(14) = 0
        RSRP1(15) = 0
        RSRP1(16) = 0
        RSRP1(17) = 0
        RSRP1(18) = 0
        RSRP1(19) = 0
        RSRP1.Update
        RSSMN.MoveNext
    Next
ElseIf CON_APP = "ROU" Then
    If RSSMN.State = 1 Then RSSMN.Close
    RSSMN.Open "SELECT ROU_IDY FROM RDY GROUP BY ROU_IDY", con, adOpenKeyset
    For A = 1 To RSSMN.RecordCount
        CTR = 0
        For dayid = 0 To 6
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT COUNT(*),ROU_IDY,DAY_IDY FROM RDY GROUP BY DAY_IDY,ROU_IDY where rou_idy='" & Trim(RSSMN!rou_idy) & "' and day_idy=" & dayid + 1, con, adOpenKeyset, adLockPessimistic
            'RS.filter = "ROU_IDY LIKE '" & Trim(RSSMN!ROU_IDY) & "' AND DAY_IDY =" & dayid + 1
            If rs.RecordCount > 0 Then CTR = CTR + (Val(rs(0)) * DAYCOUNT(dayid))
        Next
        RSRP1.AddNew
        RSRP1(0) = ""
        RSRP1(1) = ""
        RSRP1(2) = ""
        RSRP1(3) = ""
        RSRP1(4) = ""
        RSRP1(5) = ""
        RSRP1(6) = ""
        RSRP1(7) = ""
        RSRP1(8) = ""
        RSRP1(9) = Trim(RSSMN!rou_idy & "")
        RSRP1(10) = 0
        RSRP1(11) = CTR
        RSRP1(12) = 0
        RSRP1(13) = 0
        RSRP1(14) = 0
        RSRP1(15) = 0
        RSRP1(16) = 0
        RSRP1(17) = 0
        RSRP1(18) = 0
        RSRP1(19) = 0
        RSRP1.Update
        RSSMN.MoveNext
    Next
ElseIf CON_APP = "ACT" Then
    If RSSMN.State = 1 Then RSSMN.Close
    RSSMN.Open "SELECT ACT_IDY FROM ACT WHERE ACT_TYP IN('BOT','SAL') and cst_typ not in('Van') AND DBT_TYP = 'Y'", con, adOpenKeyset
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT COUNT(*),ACT_IDY,DAY_IDY FROM RDY GROUP BY DAY_IDY,ACT_IDY", con, adOpenKeyset, adLockPessimistic
    For A = 1 To RSSMN.RecordCount
        CTR = 0
        For dayid = 0 To 6
            rs.filter = "ACT_IDY LIKE '" & Trim(RSSMN!ACT_IDY) & "' AND DAY_IDY =" & dayid + 1
            If rs.RecordCount > 0 Then CTR = CTR + (Val(rs(0)) * DAYCOUNT(dayid))
        Next
        RSRP1.AddNew
        RSRP1(0) = ""
        RSRP1(1) = ""
        RSRP1(2) = ""
        RSRP1(3) = ""
        RSRP1(4) = ""
        RSRP1(5) = ""
        RSRP1(6) = ""
        RSRP1(7) = ""
        RSRP1(8) = ""
        RSRP1(9) = RSSMN!ACT_IDY
        RSRP1(10) = 0
        RSRP1(11) = CTR
        RSRP1(12) = 0
        RSRP1(13) = 0
        RSRP1(14) = 0
        RSRP1(15) = 0
        RSRP1(16) = 0
        RSRP1(17) = 0
        RSRP1(18) = 0
        RSRP1(19) = 0
        RSRP1.Update
        RSSMN.MoveNext
    Next
ElseIf CON_APP = "DAY" Then
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT COUNT(*),DAY_IDY FROM RDY GROUP BY DAY_IDY", con, adOpenKeyset, adLockPessimistic
    For A = 1 To RSSMN.RecordCount
        RSRP1.AddNew
        RSRP1(0) = ""
        RSRP1(1) = ""
        RSRP1(2) = ""
        RSRP1(3) = ""
        RSRP1(4) = ""
        RSRP1(5) = ""
        RSRP1(6) = ""
        RSRP1(7) = ""
        RSRP1(8) = ""
        RSRP1(9) = Trim(rs!DAY_IDY & "")
        RSRP1(10) = 0
        RSRP1(11) = Val(rs(0) & "")
        RSRP1(12) = 0
        RSRP1(13) = 0
        RSRP1(14) = 0
        RSRP1(15) = 0
        RSRP1(16) = 0
        RSRP1(17) = 0
        RSRP1(18) = 0
        RSRP1(19) = 0
        RSRP1.Update
        RSSMN.MoveNext
    Next
End If
End Sub

Public Sub doPent(FrmDate As Date, ToDate As Date, SMNorROU As String)
st = SMNorROU
Dim rs As New ADODB.Recordset
Dim RP1 As New ADODB.Recordset
STR1 = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'BRD' AND DIR_IDX LIKE 'IDY'")
STR2 = Split(STR1, ",")
If RP1.State = 1 Then RP1.Close
RP1.Open "select * from RP1", con, adOpenKeyset, adLockPessimistic
For A = 1 To RP1.RecordCount
    If st = "SMN" Then
        qry = "AND SMN_IDY LIKE '" & Trim(RP1!CHR_TEN) & "'"
    ElseIf st = "ROU" Then
        qry = "AND ROU_IDY LIKE '" & Trim(RP1!CHR_TEN) & "'"
    End If
    For b = 0 To UBound(STR2)
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT COUNT(DISTINCT(CST_IDY)) FROM INL WHER INV_IDY IN(SELECT INV_IDY FROM INV,PRD WHER PRD.PRD_IDY LIKE INV.PRD_IDY AND PRD.BRD_NME LIKE '" & Trim(STR2(b)) & "' AND INV_DAT >= CTOD('" & FrmDate & "') AND INV_DAT<=CTOD('" & ToDate & "'))" & qry, con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            RP1(12 + b) = Val(rs(0) & "")
            RP1.Update
        End If
    Next
    RP1.MoveNext
Next
End Sub
Public Sub doOutletsnotBilled(FrmDate As Date, ToDate As Date, SMNorROU As String)
st = SMNorROU
Dim rs As New ADODB.Recordset
Dim RP1 As New ADODB.Recordset
If RP1.State = 1 Then RP1.Close
RP1.Open "select * from RP1", con, adOpenKeyset, adLockPessimistic
For A = 1 To RP1.RecordCount
    If st = "SMN" Then
        qry = "AND SMN_IDY LIKE '" & Trim(RP1!CHR_TEN) & "'"
    ElseIf st = "ROU" Then
        qry = "AND ROU_IDY LIKE '" & Trim(RP1!CHR_TEN) & "'"
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT COUNT(DISTINCT(ACT_IDY)) FROM RDY WHERE ACT_IDY NOT IN(SELECT CST_IDY FROM INL WHERE INV_DAT >=CTOD('" & FrmDate & "') AND INV_DAT <=CTOD('" & ToDate & "') " & qry & ")" & qry, con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        RP1("NUM_THR") = Val(rs(0) & "")
        RP1.Update
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT COUNT(DISTINCT(CST_IDY)) FROM INL WHERE INV_DAT >=CTOD('" & FrmDate & "') AND INV_DAT <=CTOD('" & ToDate & "') " & qry, con, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        RP1("NUM_FOR") = Val(rs(0) & "")
        RP1.Update
      End If
    RP1.MoveNext
Next
End Sub
Public Sub doPentSKU(FrmDate As Date, ToDate As Date, SMNorROU As String)
st = SMNorROU
Dim rs As New ADODB.Recordset
Dim RP1 As New ADODB.Recordset
STR1 = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'PRD' AND DIR_IDX LIKE 'IDY'")
STR2 = Split(STR1, ",")
If RP1.State = 1 Then RP1.Close
RP1.Open "select * from RP1", con, adOpenKeyset, adLockPessimistic
For A = 1 To RP1.RecordCount
    If st = "SMN" Then
        qry = "AND SMN_IDY LIKE '" & Trim(RP1!CHR_TEN) & "'"
    ElseIf st = "ROU" Then
        qry = "AND ROU_IDY LIKE '" & Trim(RP1!CHR_TEN) & "'"
    End If
    For b = 0 To UBound(STR2)
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT COUNT(DISTINCT(CST_IDY)) FROM INL WHER INV_IDY IN(SELECT INV_IDY FROM INV,PRD WHER PRD.PRD_IDY LIKE INV.PRD_IDY AND PRD.PRD_IDY LIKE '" & Trim(STR2(b)) & "' AND INV_DAT >= CTOD('" & FrmDate & "') AND INV_DAT<=CTOD('" & ToDate & "'))" & qry, con, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            RP1(12 + b) = Val(rs(0) & "")
            RP1.Update
        End If
    Next
    RP1.MoveNext
Next
End Sub
Public Sub stkStmt(SDATE As String, edate As String)
Dim RSINV As New ADODB.Recordset
Dim rsdir As New ADODB.Recordset
Dim rsadj As New ADODB.Recordset
Dim RSOUT As New ADODB.Recordset 'Added by Ramanesh
Dim rsVan As New ADODB.Recordset
Dim rsprd As New ADODB.Recordset
Dim rsrp2 As New ADODB.Recordset
Dim RSINL As New ADODB.Recordset
Dim FYR As Date
Dim Rcpt As Single, tQty As Single, rqty As Single, bqty As Single, lqty As Single, nld As Single
Dim ORcpt As Single, Owds As Single
If rsdir.State = 1 Then rsdir.Close
rsdir.Open "SELECT DIR_ALS FROM DIR WHERE DIR_IDY='CUR' AND DIR_IDX='FYR'", con, adOpenStatic
If Not rsdir.EOF Then FYR = rsdir(0)
If IsDate(FYR) = True Then
Else
    FYR = Date$
End If
If rsVan.State = 1 Then rsVan.Close
If rsprd.State = 1 Then rsprd.Close
con.Execute "delete from rp2"
rsprd.Open "select prd_idy,prd_nme,sub_unt from prd", con, adOpenKeyset, adLockPessimistic
If rsrp2.State = 1 Then rsrp2.Close
rsrp2.Open "select * from rp2", con, adOpenKeyset, adLockPessimistic
While Not rsprd.EOF
    ostk = 0
    Rcpt = 0
    tQty = 0
    rqty = 0
    bqty = 0
    lqty = 0
    nld = 0
    cstk = 0
    sTruck = 0
    ASTK = 0
    ORcpt = 0
    Owds = 0
    outstk = 0
    salwv = 0
    tbld = 0
    tsalwv = 0
    
    If FYR = Format(SDATE, "DD/MM/YYYY") Then
        ostk = 0
    Else
        If RSINV.State = 1 Then RSINV.Close
            RSINV.Open "select sum(tot_pcs),prd_idy from inv where prd_idy like '" & rsprd("prd_idy") & "' and inv_dat < ctod('" & CDate(SDATE) & "') and TXN_TYP in('PUR','OPS','SRN','STI') and STK_TYP = 'Salable'", con, adOpenKeyset, adLockPessimistic
        
        If RSOUT.State = 1 Then RSOUT.Close
            RSOUT.Open "select sum(tot_pcs),prd_idy from inv where prd_idy like '" & rsprd("prd_idy") & "' and inv_dat < ctod('" & CDate(SDATE) & "') and TXN_TYP in('STO','PRT') and STK_TYP = 'Salable'", con, adOpenKeyset, adLockPessimistic
            
        If Not RSOUT.EOF Then outstk = (RSOUT(0))
        
        If rsVan.State = 1 Then rsVan.Close
            rsVan.Open "select (sum(int(TOT_QTY))* SUB_UNT)+((sum(TOT_QTY)-sum(int(TOT_QTY)))*1000),(sum(int(RET_QTY))* SUB_UNT)+((sum(RET_QTY)-sum(int(RET_QTY)))*1000),(sum(int(BRK_QTY))* SUB_UNT)+((sum(BRK_QTY)-sum(int(BRK_QTY)))*1000),(sum(int(LKG_QTY))* SUB_UNT)+((sum(LKG_QTY)-sum(int(LKG_QTY)))*1000) from VAN, PRD where VAN.PRD_IDY like '" & rsprd("prd_idy") & "' and (van_cls='S' or van_cls='C' or van_cls = 'N') and VAN.PRD_IDY = PRD.PRD_IDY and inv_dat < ctod('" & CDate(SDATE) & "')", con, adOpenKeyset, adLockPessimistic
        
        If Not RSINV.EOF Then Rcpt = Val(RSINV(0))
            Rcpt = Rcpt - outstk
            
        If RSINL.State = 1 Then RSINL.Close
            RSINL.Open "select sum(tot_pcs) from inv where (ord_idy is null or ord_idy='' or ord_idy=' ') and prd_idy like '" & rsprd("prd_idy") & "' and inv_dat < ctod('" & CDate(SDATE) & "') and txn_typ='SAL' and stk_typ in('Salable','Free')", con, adOpenKeyset, adLockPessimistic
        
        If Not RSINL.EOF Then salwv = Val(Trim(RSINL(0)) & "")
        
        If RSINL.State = 1 Then RSINL.Close
            RSINL.Open "select sum(tot_pcs) from inv where (ord_idy is null or ord_idy='' or ord_idy=' ') and prd_idy like '" & rsprd("prd_idy") & "' and inv_dat < ctod('" & CDate(SDATE) & "') and txn_typ='SAL' and stk_typ in('TakeBack','Breakage','Leakage','Damage')", con, adOpenKeyset, adLockPessimistic
            
        If Not RSINL.EOF Then tbld = Val(Trim(RSINL(0)) & "")
            tsalwv = salwv - tbld
            If Not rsVan.EOF Then
                tQty = rsVan(0)
                rqty = rsVan(1)
                bqty = rsVan(2)
                lqty = rsVan(3)
                nld = tQty - (rqty + bqty + lqty)
            End If
            ostk = Rcpt - (nld + tsalwv)  ' Opening Stock
    End If
    ' Adjustments Amount Adding
        If rsadj.State = 1 Then rsadj.Close
        rsadj.Open "select prd_qty,prd_idy,opn_qty,iss_qty from inv where prd_idy like '" & rsprd("prd_idy") & "' and inv_dat>=ctod('" & FYR & "') and inv_dat<=ctod('" & CDate(SDATE) - 1 & "') and txn_typ='ADJ' and stk_typ='Salable'", con, adOpenKeyset, adLockPessimistic
        While Not rsadj.EOF
            If Val(rsadj(0)) > 0 Then
                ostk = ostk + MTS(Val(rsadj(0)), Val(rsprd(2)))
            Else
                ostk = ostk - MTS(Val(Abs(rsadj(0))), Val(rsprd(2)))
            End If
            rsadj.MoveNext
        Wend
    ' Adjustments End
    
    Rcpt = 0
    nld = 0
    salwv = 0
    tbld = 0
    tsalwv = 0
    If RSINV.State = 1 Then RSINV.Close
    
    RSINV.Open "select sum(tot_pcs),prd_idy from inv where prd_idy like '" & rsprd("prd_idy") & "' and inv_dat>=ctod('" & CDate(SDATE) & "') and inv_dat<=ctod('" & CDate(edate) & "') and txn_typ IN ('PUR') and stk_typ='Salable'", con, adOpenKeyset, adLockPessimistic
    
    If rsVan.State = 1 Then rsVan.Close
        rsVan.Open "select (sum(int(TOT_QTY))* SUB_UNT)+((sum(TOT_QTY)-sum(int(TOT_QTY)))*1000),(sum(int(RET_QTY))* SUB_UNT)+((sum(RET_QTY)-sum(int(RET_QTY)))*1000),(sum(int(BRK_QTY))* SUB_UNT)+((sum(BRK_QTY)-sum(int(BRK_QTY)))*1000),(sum(int(LKG_QTY))* SUB_UNT)+((sum(LKG_QTY)-sum(int(LKG_QTY)))*1000) from VAN, PRD where VAN.PRD_IDY like '" & rsprd("prd_idy") & "' and (van_cls='S' or van_cls='C' or van_cls = 'N') and VAN.PRD_IDY = PRD.PRD_IDY and inv_dat>=ctod('" & CDate(SDATE) & "') and inv_dat<=ctod('" & CDate(edate) & "')", con, adOpenKeyset, adLockPessimistic
        
    If Not RSINV.EOF Then Rcpt = (RSINV(0))
    
    If RSINL.State = 1 Then RSINL.Close
        RSINL.Open "select sum(tot_pcs) from inv where (ord_idy is null or ord_idy='' or ord_idy=' ') and prd_idy like '" & rsprd("prd_idy") & "' and inv_dat>=ctod('" & CDate(SDATE) & "') and inv_dat<=ctod('" & CDate(edate) & "') and txn_typ='SAL' and stk_typ in('Salable','Free')", con, adOpenKeyset, adLockPessimistic
        
    If Not RSINL.EOF Then salwv = Val(Trim(RSINL(0)) & "")
    
    If RSINL.State = 1 Then RSINL.Close
        RSINL.Open "select sum(tot_pcs) from inv where (ord_idy is null or ord_idy='' or ord_idy=' ') and prd_idy like '" & rsprd("prd_idy") & "' and inv_dat>=ctod('" & CDate(SDATE) & "') and inv_dat<=ctod('" & CDate(edate) & "') and txn_typ='SAL' and stk_typ in('TakeBack','Breakage','Leakage','Damage')", con, adOpenKeyset, adLockPessimistic
        If Not RSINL.EOF Then tbld = Val(Trim(RSINL(0)) & "")
        
        tsalwv = salwv - tbld
        
    If Not rsVan.EOF Then
        tQty = rsVan(0)
        rqty = rsVan(1)
        bqty = rsVan(2)
        lqty = rsVan(3)
        nld = tQty - (rqty + bqty + lqty)
    End If
    
    If rsVan.State = 1 Then rsVan.Close
        rsVan.Open "select (sum(int(TOT_QTY))* SUB_UNT)+((sum(TOT_QTY)-sum(int(TOT_QTY)))*1000) from VAN, PRD where VAN.PRD_IDY like '" & rsprd("prd_idy") & "' and VAN.PRD_IDY = PRD.PRD_IDY and VAN_CLS = 'N' and inv_dat>=ctod('" & CDate(SDATE) & "') and inv_dat<=ctod('" & CDate(edate) & "')", con, adOpenKeyset, adLockPessimistic
    If Not rsVan.EOF Then sTruck = rsVan(0)
    
    If RSINV.State = 1 Then RSINV.Close
        RSINV.Open "SELECT SUM(TOT_PCS) FROM INV WHERE prd_idy like '" & rsprd("prd_idy") & "' and inv_dat>=ctod('" & CDate(SDATE) & "') and inv_dat<=ctod('" & CDate(edate) & "') and txn_typ in('SRN','STI','OPS') and stk_typ='Salable'", con, adOpenKeyset, adLockPessimistic
    
    If Not RSINV.EOF Then ORcpt = RSINV(0)
    
    If RSINV.State = 1 Then RSINV.Close
        RSINV.Open "SELECT SUM(TOT_PCS) FROM INV WHERE prd_idy like '" & rsprd("prd_idy") & "' and inv_dat>=ctod('" & CDate(SDATE) & "') and inv_dat<=ctod('" & CDate(edate) & "') and txn_typ in('PRT','STO') and stk_typ='Salable'", con, adOpenKeyset, adLockPessimistic
        
    If Not RSINV.EOF Then Owds = RSINV(0)
        'cstk = (ostk + Rcpt + ORcpt) - (sTruck + nld + Owds)
        cstk = (ostk + Rcpt + ORcpt) - (nld + Owds) 'Formula corrected as per TSR's instructions by PNR
    
    If RSINV.State = 1 Then RSINV.Close
        RSINV.Open "select TOT_PCS,PRD_IDY,OPN_QTY,ISS_QTY from INV where PRD_IDY like '" & rsprd("prd_idy") & "' and inv_dat>=ctod('" & CDate(SDATE) & "') and inv_dat<=ctod('" & CDate(edate) & "') and txn_typ='ADJ' and stk_typ='Salable'", con, adOpenKeyset, adLockPessimistic
        
    While Not RSINV.EOF
        If Val(RSINV(0)) > 0 Then
            cstk = cstk + RSINV(0)
            ASTK = ASTK + RSINV(0)
        Else
            cstk = cstk - RSINV(0)
            ASTK = ASTK - RSINV(0)
        End If
        RSINV.MoveNext
    Wend
    
    rsrp2.AddNew
    rsrp2(0) = rsprd(0)
    rsrp2(10) = STM(Val(ostk), Val(rsprd(2)))
    rsrp2(11) = STM(Val(Rcpt), Val(rsprd(2)))
    rsrp2(12) = STM(Val(ORcpt), Val(rsprd(2)))
    rsrp2(13) = STM(Val(sTruck), Val(rsprd(2)))
    rsrp2(14) = STM(Val(nld + tsalwv), Val(rsprd(2)))
    rsrp2(15) = STM(Val(ASTK), Val(rsprd(2)))
    rsrp2(16) = STM(Val(Owds), Val(rsprd(2)))
    rsrp2(17) = STM(Val(cstk - tsalwv), Val(rsprd(2)))
    rsrp2.Update
    rsprd.MoveNext
Wend
'OpenReport CURDIR & "\VVR_RPT.RP1", "Stock Summary Net Load", 2, "PRD->STK_TYP <> 'Closed'", 6
OpenReport CURDIR & "\VVR_RPT.RP1", "Stock Summary Net Load CasePKT", 2, "PRD->STK_TYP <> 'Closed'", 6
End Sub
Public Function AMT1(amount As String) As String
A = InStr(amount, ".")
If Len(amount) = 1 And A = 1 Then
     AMT1 = "0.00"
     Exit Function
End If
If A > 1 Then
    If InStr(A + 1, amount, ".") > 0 Then
        AMT1 = "0.00"
    Else
        AMT1 = Mid(amount, 1, A) & Mid(amount, A + 1, 2)
    End If
ElseIf A = 1 And Len(amount) = 2 Then
    AMT1 = "0" & amount & "0"
ElseIf A = 1 And Len(amount) >= 3 Then
    AMT1 = "0" & Mid(amount, 1, 3)
ElseIf A = 0 And Len(Trim(amount)) >= 7 Then
    AMT1 = Trim(amount) & ".00"
ElseIf Len(Trim(amount)) = 7 Then
    AMT1 = Mid(amount, 1, 6) & "." & Mid(amount, 7, 1) & "0"
ElseIf Len(Trim(amount)) = 0 Then
   AMT1 = "0.00"
ElseIf Len(Trim(amount)) < 7 Then
    AMT1 = amount & ".00"
End If
End Function

Public Function convert1(num As String) As String
If Len(Trim(num)) >= 10 Then
    convert1 = num
Else
    num = Round(Val(num), 2)
    If InStr(num, ".") > 0 Then
        r = Mid(num, InStr(num, ".") + 1, Len(num) - InStr(num, "."))
        If Trim(Len(r)) = 1 Then
            convert1 = Mid(num, 1, InStr(num, ".") - 1) & "." & r & "0"
        ElseIf Trim(Len(r)) = 2 Then
            convert1 = num
        End If
    Else
        If Trim(Len(num)) <> 0 Then
            convert1 = num & ".00"
        Else
            convert1 = "0.00"
        End If
        End If
End If
End Function

Public Function preincrement(FLDVAL As String) As String
Do While Len(icrt) < 5 - Len(FLDVAL)
    icrt = icrt & "0"
Loop
preincrement = Trim(icrt & FLDVAL)
End Function
Public Sub EMP_SHL(SHL_TAG As String)
M_FRM_PRV.Tag = SHL_TAG
M_FRM_PRV.Check3.Caption = "Check for Empties Module after Save"
M_FRM_PRV.Caption = "Empties"
If check("prt", "prt_nme", SHL_TAG) = True Then
    M_FRM_PRV.Check3.Value = 1
Else
    M_FRM_PRV.Check3.Value = 0
End If
M_FRM_PRV.Show 1
End Sub
Public Sub LOADSHELLS(TXN_TYP As String, invid As String)
If MsgBox("Change Empties Details ?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
    Gltxninvid = TXN_TYP & "," & invid
    M_FRM_TRS.Show 1
End If
End Sub
Public Sub LOADSHELLSSORTING(TXN_TYP As String, invid As String)
If MsgBox("Sort Empties and Print EDA ?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
    Gltxninvid = TXN_TYP & "," & invid
    M_FRM_SSI.Show 1
End If
End Sub
Public Sub EMP_SHL_SOR(SHL_TAG As String)
M_FRM_PRV.Tag = SHL_TAG
M_FRM_PRV.Check3.Caption = "Check for Empties Sort After Save"
M_FRM_PRV.Caption = "Sort Empties"
If check("prt", "prt_nme", SHL_TAG) = True Then
    M_FRM_PRV.Check3.Value = 1
Else
    M_FRM_PRV.Check3.Value = 0
End If
M_FRM_PRV.Show 1
End Sub
Public Sub IMP_BKP()
Dim FS As Object, f As Object, F1 As Object, FN As Object
Dim cn As New ADODB.Connection
Dim CHK_DIR As Integer
CHK_DIR = CHK_IMP_BKP()

If Dir(CURDIR & "\bkp" & CHK_DIR, vbDirectory) = "" Then


    MkDir CURDIR & "\bkp" & CHK_DIR
    Destdir = CURDIR & "\bkp" & CHK_DIR & "\"
    Set FS = CreateObject("Scripting.FileSystemObject")
       
    SPATH = CURDIR & "\???.dbf"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    SPATH = CURDIR & "\eff.*"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    SPATH = CURDIR & "\*.cdx"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    SPATH = CURDIR & "\*.RP1"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    SPATH = CURDIR & "\*.fpt"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    SPATH = CURDIR & "\*.RRW"
    DPATH = Destdir
    FS.CopyFile SPATH, DPATH
    CopyFile CURDIR & "\IND.EXE", Destdir & "IND.EXE", 0
    CopyFile CURDIR & "\CND.EXE", Destdir & "CND.EXE", 0
    CopyFile CURDIR & "\RRUNTIME.EXE", Destdir & "RRUNTIME.EXE", 0
    CopyFile CURDIR & "\RRUN.EXE", Destdir & "RRUN.EXE", 0
    CopyFile CURDIR & "\RR.CNF", Destdir & "RR.CNF", 0
    CopyFile CURDIR & "\PRI.EXE", Destdir & "PRI.EXE", 0
    CopyFile CURDIR & "\PRB.EXE", Destdir & "PRB.EXE", 0
  End If
End Sub
Public Sub TxnMoveBackup(SRC_DB As String, TGT_DB As String, B_TXN_IDY As String, B_TXN_TYP As String)
CHK_TXT = Trim(SRC_DB) & ".DBF," & Trim(TGT_DB) & ".DBF"
If CHECKTB(Trim(CHK_TXT)) = 1 Then
    MsgBox "Data Integrity Loss! Report Error for Support - Help!", vbCritical, "DAS Version 4.0"
    Exit Sub
End If
'Code by PNR for DataBackup - Deleted Recordset before Deletion for Critical Tables in Schema
'Restore Manager Can Restore from these tables upon Sync Verification
Dim SRC_RSET As New ADODB.Recordset
Dim TGT_RSET As New ADODB.Recordset

If TGT_RSET.State = 1 Then TGT_RSET.Close
TGT_RSET.Open "SELECT * FROM " & Trim(TGT_DB) & " where INV_IDY = ''", con, adOpenKeyset, adLockPessimistic

If SRC_RSET.State = 1 Then SRC_RSET.Close
'If Trim(B_TXN_TYP) = "" Then
    If Trim(SRC_DB) = "LNV" Or Trim(SRC_DB) = "CNV" Then
        SRC_RSET.Open "SELECT DISTINCT * FROM " & Trim(SRC_DB) & " where INV_IDY like '" & Trim(B_TXN_IDY) & "' GROUP BY LEV_DES", con, adOpenKeyset, adLockPessimistic
    End If
    If Trim(SRC_DB) = "VAN" Or Trim(SRC_DB) = "VBK" Then
        SRC_RSET.Open "SELECT DISTINCT * FROM " & Trim(SRC_DB) & " where INV_IDY like '" & Trim(B_TXN_IDY) & "' GROUP BY BAT_IDY", con, adOpenKeyset, adLockPessimistic
    End If
    If Trim(SRC_DB) = "SNV" Or Trim(SRC_DB) = "SBK" Then
        If SRC_RSET.State = 1 Then SRC_RSET.Close
        SRC_RSET.Open "SELECT DISTINCT * FROM " & Trim(SRC_DB) & " where TXN_IDY like '" & Trim(B_TXN_IDY) & "' GROUP BY PRD_PCK", con, adOpenKeyset, adLockPessimistic
    End If
'Else
    If Trim(SRC_DB) = "INV" Or Trim(SRC_DB) = "PNV" Or Trim(SRC_DB) = "XNV" Then
        SRC_RSET.Open "SELECT DISTINCT * FROM " & Trim(SRC_DB) & " where INV_IDY like '" & Trim(B_TXN_IDY) & "' AND TXN_TYP like '" & Trim(B_TXN_TYP) & "' GROUP BY BAT_IDY", con, adOpenKeyset, adLockPessimistic
    End If
    If Trim(SRC_DB) = "INL" Or Trim(SRC_DB) = "PNL" Or Trim(SRC_DB) = "XNL" Then
        SRC_RSET.Open "SELECT DISTINCT * FROM " & Trim(SRC_DB) & " where INV_IDY like '" & Trim(B_TXN_IDY) & "' AND TXN_TYP like '" & Trim(B_TXN_TYP) & "' GROUP BY INV_IDY", con, adOpenKeyset, adLockPessimistic
    End If
'End If
If SRC_RSET.State = 1 Then
While Not SRC_RSET.EOF
    TGT_RSET.AddNew
    For K = 0 To TGT_RSET.Fields.count - 1
        If TGT_RSET(K).Type = adVarChar Or TGT_RSET(K).Type = adBSTR Or TGT_RSET(K).Type = adChar Or TGT_RSET(K).Type = adLongVarChar Or TGT_RSET(K).Type = adSmallInt Then
            TGT_RSET(K) = ""
        ElseIf TGT_RSET(K).Type = adInteger Or TGT_RSET(K).Type = adNumeric Or TGT_RSET(K).Type = adBigInt Or TGT_RSET(K).Type = adSingle Or TGT_RSET(K).Type = adDouble Then
            TGT_RSET(K) = 0
        ElseIf TGT_RSET(K).Type = adDate Or TGT_RSET(K).Type = adDBDate Or TGT_RSET(K).Type = adDBTimeStamp Then
            TGT_RSET(K) = Date
        End If
    Next
    For K = 0 To SRC_RSET.Fields.count - 1
        FLD = SRC_RSET.Fields(K).name
        If Trim(SRC_RSET(FLD)) = "" Then
            TGT_RSET(FLD) = ""
        Else
            TGT_RSET(FLD) = SRC_RSET(FLD)
        End If
    Next
    TGT_RSET.Update
    SRC_RSET.MoveNext
Wend
End If
Set TGT_RSET = Nothing
Set SRC_RSET = Nothing
End Sub
Public Function DongleResult(dResult) As String
    Select Case dResult
        
        Case 1
        DongleResult = "1 (Success)"
        
        Case 0
        DongleResult = "0 (Error: Different Dongle)"
        
        Case -1
        DongleResult = "-1 (Error: Dongle Missing)"
        
        Case -2
        DongleResult = "-2 (Error: OS found is NT/2000 and Driver COPYLOCK.SYS not found loaded)"
        
        Case -3
        DongleResult = "-3 (Error: Invalid Password)"
        
        Case -4
        DongleResult = "-4 (Error: Facility N.A. with this model of Dongle)"
        
        Case -5
        DongleResult = "-5 (Error: Can not write into Dongle)"
        
        Case -11
        DongleResult = "-11 (Invalid API library file (may be corrupted))"
        
        Case -6
        DongleResult = "-6 ([NetCopylock]Error: Login Count exhausted)"
        
        Case -7
        DongleResult = "-7 ([NetCopylock]Error: can not Logout)"
        
        Case -8
        DongleResult = "-8 ([NetCopylock]Error: can not open communication channel)"
        
        Case -9
        DongleResult = "-9 ([NetCopylock]Error: can not connect to NetDongle server)"
        
        Case -10
        DongleResult = "-10([NetCopylock]Error: can not close communication channel)"
        
    End Select
End Function

''''''                        While Val(AVB_SCH_QTY) = 0
'''''                            REC_SET.AddNew
'''''                            REC_SET("inv_idy") = Trim(INC)
'''''                            REC_SET("inv_dat") = Trim(M_INV_DAT)
'''''                            REC_SET("prd_idy") = .TextMatrix(A, 5)
'''''                            REC_SET("LST_RAT") = GetValue("LST_RAT", "PRD WHERE PRD_IDY LIKE '" & MSS.TextMatrix(A, 5) & "'")
'''''                            REC_SET("LST_DIS") = GetValue("LST_DIS", "PRD WHERE PRD_IDY LIKE '" & MSS.TextMatrix(A, 5) & "'")
'''''                            REC_SET("STK_TYP") = "Free"
'''''                            REC_SET("PRD_QTY") = STM(AVB_SCH_QTY, Val(.TextMatrix(A, 12))) 'STM(Val(.TextMatrix(A, 2)), Val(.TextMatrix(A, 12)))
'''''                            REC_SET("prd_rat") = Val(.TextMatrix(A, 3))
'''''                            REC_SET("txn_typ") = "SAL"
'''''                            REC_SET!BAT_IDY = Trim(.TextMatrix(A, 7))
'''''                            REC_SET!SUB_UNT = Val(.TextMatrix(A, 12))
'''''                            REC_SET!PRD_UNT = STM(Val(AVB_SCH_QTY), Val(.TextMatrix(A, 12))) 'Int(STM(Val(.TextMatrix(A, 2)), Val(.TextMatrix(A, 12))))
'''''                            REC_SET!PRD_PCS = Val(AVB_SCH_QTY) Mod Val(.TextMatrix(A, 12)) 'Val(.TextMatrix(A, 2)) Mod Val(.TextMatrix(A, 12))
'''''                            REC_SET!TOT_PCS = Val(AVB_SCH_QTY) 'Val(.TextMatrix(A, 2))
'''''                            REC_SET!PRD_UOM = Round(Val(.TextMatrix(A, 16)), 2)
'''''                            REC_SET!Sch_Idy = Trim(.TextMatrix(A, 8))
'''''                            REC_SET("COL_AMT") = Val((.TextMatrix(A, 4)))
'''''                            REC_SET!ORD_IDY = M_VAN_IDY
'''''                            If Len(Trim(M_VAN_IDY)) = 0 Then
'''''                                REC_SET!stk_upd = "Y"
'''''                            Else
'''''                                REC_SET!stk_upd = "N"
'''''                            End If
'''''                            REC_SET.Update
'''''                            CON.Execute "insert into lnv values('" & Trim(INC) & "','" & Trim(.TextMatrix(A, 8)) & "','" & Trim(.TextMatrix(A, 7)) & "'," & Val(.TextMatrix(A, 2)) & "," & Val((.TextMatrix(A, 4))) & ",0,'FRE')"
'''''                            If Len(Trim(M_VAN_IDY)) = 0 Then
'''''                                CON.Execute "update PRD where prd_idy like '" & Trim(.TextMatrix(A, 5)) & "' set CUR_LEV = " & CONQTY(QTY(STM(.TextMatrix(A, 2), Val(.TextMatrix(A, 12)))), Trim(.TextMatrix(A, 5))), d
'''''                                CON.Execute "UPDATE PRD WHERE PRD_IDY LIKE '" & Trim(.TextMatrix(A, 5)) & "' set PRD_PCS = PRD_PCS - " & Val(.TextMatrix(A, 2)), d
'''''                                CON.Execute "UPDATE BRD WHERE BAT_IDY LIKE '" & Trim(.TextMatrix(A, 7)) & "' set prd_qty = " & CONBRDQTY(STM(.TextMatrix(A, 2), Val(.TextMatrix(A, 12))), Trim(.TextMatrix(A, 7))), d
'''''                            End If
'''''                          '  ESC_QTY = Val(ESC_QTY) + Val(.TextMatrix(A, 2))
'''''                            AVB_SCH_QTY = Val(AVB_SCH_QTY) - Val(.TextMatrix(A, 2))
''''''                        Wend



