Attribute VB_Name = "EFF_LIB1"
Option Compare Text
Public osversion As Integer
Public PrdCollection As New Collection
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
Public groupID, ACCOUNTTYP As String
Public CompanyName As String
Public CompanyFolderPath As String

Public DEL_REC_IDY As String
Public sprg As String
Public SAL_QRY As String
Public RN As Integer
Public CNU As Integer
Public cn As New ADODB.Connection
Public CON As New ADODB.Connection
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
Public UserName As String
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
'Public CN As New ADODB.Connection
Public CN1 As New ADODB.Connection
Public cn2 As New ADODB.Connection
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public DB As New ADODB.Connection
Public rs As New ADODB.Recordset
Public SRC_RSET As New ADODB.Recordset
Public TGT_RSET As New ADODB.Recordset
Public TB As TableDef
Public RS1 As New ADODB.Recordset
Public tb1 As TableDef
Public dbs As New ADODB.Connection
Public TNAME As String
Dim strPath As String
Public Gltxninvid As String
Public Myzip As ColaZip.GSColaZip32
Public Sub AddData(ByRef ctrlName As Control, frmName As Form, fldName As String, tblName As String, Condition As String, Group_order As String)
Dim rs As New ADODB.Recordset
rs.Open "select " & Trim(fldName) & " from " & Trim(tblName) & " " & Condition & " " & Group_order, CON, adOpenKeyset, adLockPessimistic
ctrlName.CLEAR
ctrlName.AddItem ""
While Not rs.EOF
    If Trim(rs(0) & "") <> "" Then ctrlName.AddItem Trim(rs(0) & "")
    rs.MoveNext
Wend
End Sub

Public Function CHK_IMP_BKP() As Integer
Dim FS As Object, f As Object, F1 As Object, FN As Object
 Dim BKP_INT As Integer
 BKP_INT = 0
 Set FS = CreateObject("Scripting.FileSystemObject")
 Set f = FS.GetFolder(RootPath)
 Set FN = f.SubFolders
 If FN.COUNT > 0 Then
    For Each F1 In FN
    If Mid(F1, Len(RootPath) + 2, 3) = "BKP" Then
        If Val(BKP_INT) < Val(Mid(F1, Len(RootPath) + 2 + 3)) Then
            BKP_INT = Val(Mid(F1, Len(RootPath) + 2 + 3))
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

Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
CON.Execute "DELETE FROM SCL"
RS1.Open "SELECT * FROM SCL ", CON, adOpenKeyset, adLockPessimistic
rs.Open "SELECT SUM(R_SHL_NOS+DP_SHL_NOS+BRK_SHL),SUM(R_PCK_NOS+DP_PCK_NOS+BRK_PCK),CST_IDY,SHL_IDY,TXN_TYP FROM SNV WHERE TXN_TYP IN('SAL','PRT','STO','AIN') GROUP BY CST_IDY,SHL_IDY", CON, adOpenKeyset, adLockPessimistic
While Not rs.EOF = True
    RS1.AddNew
    RS1!INV_DAT = Date
    RS1!Sch_Idy = rs!shl_idy
    RS1!prd_idy = rs!CST_IDY
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
rs.Open "SELECT SUM(PRD_UNT),SUM(TOT_PCS),SHL_IDY,CST_IDY,INV.TXN_TYP FROM INV,SHL,INL WHERE INL.INV_IDY = INV.INV_IDY AND SHL.PRD_IDY = INV.PRD_IDY AND SHL_TYP='Shell' AND  INL.STK_TYP  IN('Salable','Free') GROUP BY CST_IDY,SHL_IDY,INV.TXN_TYP", CON, adOpenKeyset, adLockPessimistic
While Not rs.EOF
    If rs!TXN_TYP = "SAL" Or rs!TXN_TYP = "STO" Or rs!TXN_TYP = "PRT" Then
        CON.Execute "update scl set OPN_STK=" & Val(rs(0)) & " where sch_idy = '" & Trim(rs!shl_idy & "") & "' and prd_idy = '" & Trim(rs!CST_IDY) & "'", d
        CON.Execute "update scl set RCT_STK=" & Val(rs(1)) & " where sch_idy = '" & Trim(rs!shl_idy & "") & "' and prd_idy = '" & Trim(rs!CST_IDY) & "'", d
        If d = 0 Then
            RS1.AddNew
            RS1!INV_DAT = Date
            RS1!Sch_Idy = rs!shl_idy
            RS1!prd_idy = rs!CST_IDY
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
        CON.Execute "update scl set ISS_STK=" & Val(rs(0)) & " where sch_idy = '" & Trim(rs!shl_idy & "") & "' and prd_idy = '" & Trim(rs!CST_IDY) & "'", d
        CON.Execute "update scl set CLS_STK=" & Val(rs(1)) & " where sch_idy = '" & Trim(rs!shl_idy & "") & "' and prd_idy = '" & Trim(rs!CST_IDY) & "'", d
        If d = 0 Then
            RS1.AddNew
            RS1!INV_DAT = Date
            RS1!Sch_Idy = rs!shl_idy
            RS1!prd_idy = rs!CST_IDY
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
    a = a + 1
    rs.MoveNext
Wend
If rs.State = 1 Then rs.Close
rs.Open "SELECT SUM(R_SHL_NOS+DP_SHL_NOS+BRK_SHL),SUM(R_PCK_NOS+DP_PCK_NOS+BRK_PCK),CST_IDY,SHL_IDY,TXN_TYP FROM SNV WHERE TXN_TYP IN('RCT','PUR') GROUP BY CST_IDY,SHL_IDY", CON, adOpenKeyset, adLockPessimistic
While Not rs.EOF
    CON.Execute "update scl set OPN_STK=" & Val(rs(0)) & " where sch_idy = '" & Trim(rs!shl_idy & "") & "' and prd_idy = '" & Trim(rs!CST_IDY) & "'", d
    CON.Execute "update scl set RCT_STK=" & Val(rs(1)) & " where sch_idy = '" & Trim(rs!shl_idy & "") & "' and prd_idy = '" & Trim(rs!CST_IDY) & "'", d
    If d = 0 Then
        RS1.AddNew
        RS1!INV_DAT = Date
        RS1!Sch_Idy = rs!shl_idy
        RS1!prd_idy = rs!CST_IDY
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
CON.Execute "update scl set sal_stk= opn_stk-iss_stk"
CON.Execute "update scl set fre_stk= rct_stk-cls_stk"
End Sub
Public Function exitform(FRM As Form)
'FRM.Command1.SetFocus
End Function

'Public Sub clrctr(FRM As Form)
'For a = 1 To FRM.count - 1
'If TypeOf FRM.Controls(a) Is TextBox Then FRM.Controls(a).Text = ""
'If TypeOf FRM.Controls(a) Is MaskEdBox Then FRM.Controls(a).Text = ""
'Next
'End Sub

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
    For I = i1 + 1 To 10
    convert11 = " " & convert11
    Next
End If

End Function

Public Function check(tbnme As String, fldnme As String, fldvalue As String) As Boolean

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from " & tbnme & " where " & fldnme & " = '" & Trim(fldvalue) & "'", CON, adOpenStatic
If rs.RecordCount > 0 Then
check = True
Else
check = False
End If
If rs.State = 1 Then rs.Close
End Function
Public Function querycheck(tbnme As String, fldnme As String, fldvalue As String, qry As String) As Boolean


Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from " & tbnme & " where " & fldnme & " = '" & Trim(fldvalue) & "' " & Trim(qry), CON, adOpenStatic
If rs.RecordCount > 0 Then
querycheck = True
Else
querycheck = False
End If
If rs.State = 1 Then rs.Close
End Function

Public Function Check1(tbnme As String, fldnme As String, fldvalue As String, TYP As String) As Boolean

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from " & tbnme & " where " & fldnme & " = '" & Trim(fldvalue) & "' ", CON, adOpenStatic
If rs.RecordCount > 0 Then
Check1 = True
Else
Check1 = False
End If
If rs.State = 1 Then rs.Close
End Function

Public Function increment(FLDVAL As String) As String

    Dim ict As Integer
    Dim icr As Integer
    Dim icrt As String
    Dim rec As New ADODB.Recordset
    If rec.State = 1 Then rec.Close
    rec.Open "select " & FLDVAL & " from idy", CON, adOpenKeyset, adLockPessimistic
    If rec.EOF = True Then rec.AddNew
    rec(0) = Val(Trim(rec(0)) & "") + 1
    ict = Val(Trim(rec(0)) & "")
    rec.Update
    If rec.State = 1 Then rec.Close
        
    rec.Open "select pre_" & Mid(FLDVAL, 5, 3) & " from pre", CON, adOpenKeyset, adLockPessimistic
    If rec.EOF = False Then
        icrt = Trim(rec(0) & "")
    End If
    
    Do While Len(icrt) <= 10 - Len(STR(ict))
        icrt = icrt & "0"
    Loop
    If rec.State = 1 Then rec.Close
    increment = icrt & ict
End Function

Public Function T7increment(FLDVAL As String) As String


Dim ict, icr As Integer, icrt As String, rec As New ADODB.Recordset
If rec.State = 1 Then rec.Close
rec.Open "select " & FLDVAL & " from idy", CON, adOpenKeyset, adLockPessimistic
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


Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select " & FLDVAL & " from idy", CON, adOpenKeyset, adLockPessimistic
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
a = InStr(dis, ".")
b = InStr(dis, "-")
If InStr(a + 1, dis, ".") > 0 Or (Len(Trim(dis)) = 1 And a = 1) Then
    discount = "0.00"
    Exit Function
End If
If b = 1 And Len(dis) = 1 Then
    discount = "0.00"
    Exit Function
ElseIf b > 1 And Len(dis) >= 2 Then
    discount = "0.00"
    Exit Function
ElseIf b = 1 And Len(dis) >= 5 And a = 0 Then
    discount = Mid(dis, 1, 3) & "." & Mid(dis, 4, 2)
    Exit Function
ElseIf b = 1 And Len(dis) >= 5 And a = 4 Then
    discount = dis
    Exit Function
ElseIf b = 1 And Len(dis) = 3 And a = 0 Then
    discount = dis & ".00"
    Exit Function
ElseIf b = 1 And Len(dis) >= 3 And a = 2 Then
    discount = Mid(dis, 1, 1) & "0" & Mid(dis, 2, 3)
    Exit Function
    
ElseIf b = 1 And Len(dis) = 2 And a = 2 Then
    discount = "0.00"
    Exit Function
ElseIf b = 1 And Len(dis) = 5 And a = 5 Then
    discount = Mid(dis, 1, 3) & "." & Mid(dis, 4, 1) & "0"
    Exit Function
ElseIf b = 1 And Len(dis) > 5 And a = 5 Then
    discount = Mid(dis, 1, 3) & "." & Mid(dis, 4, 1) & Mid(dis, 6, 1)
    Exit Function
ElseIf b = 1 And Len(dis) > 5 And a = 6 Then
    discount = Mid(dis, 1, 3) & "." & Mid(dis, 4, 2)
    Exit Function
End If
If Len(dis) >= 4 And a = 0 Then
    discount = Mid(dis, 1, 2) & "." & Mid(dis, 3, 2)
ElseIf Len(dis) > 4 And a = 4 Then
    discount = Mid(dis, 1, 2) & "." & Mid(dis, 3, 1) & Mid(dis, 5, 1)
ElseIf Len(dis) = 4 And a = 4 Then
    discount = Mid(dis, 1, 2) & "." & Mid(dis, 3, 1) & "0"
ElseIf Len(dis) >= 4 And a = 2 Then
    discount = Mid(dis, 1, 2) & Mid(dis, 3, 2)
ElseIf Len(dis) = 4 And a = 3 Then
    discount = Mid(dis, 1, 3) & Mid(dis, 4, 1) & "0"
ElseIf Len(dis) = 6 And a = 3 Then
    discount = Mid(dis, 1, 3) & Mid(dis, 4, 2)
ElseIf Len(dis) >= 4 And (a = 5 Or a = 6) Then
    discount = Mid(dis, 1, 2) & "." & Mid(dis, 3, 2)
ElseIf Mid(dis, 1, 1) = "." And Mid(dis, 2, 1) = "." Then
    discount = "0.00"
ElseIf Len(dis) >= 3 And a = 1 Then
    discount = "0" & Mid(dis, 1, 3)
ElseIf Len(dis) = 3 And a = 0 Then
    discount = Mid(dis, 1, 2) & "." & Mid(dis, 3, 1) & "0"
ElseIf Len(dis) = 3 And a = 2 Then
    discount = Mid(dis, 1, 3) & "0"
ElseIf Len(dis) = 3 And a = 3 Then
    discount = Mid(dis, 1, 3) & "00"
ElseIf Len(dis) = 2 And a = 1 Then
    discount = "0" & Mid(dis, 1, 2) & "0"
ElseIf Len(dis) = 2 And a = 2 Then
    discount = Mid(dis, 1, 2) & "00"
ElseIf (Len(dis) = 1 Or Len(dis) = 2) And a = 0 Then
    discount = dis & ".00"
ElseIf Len(dis) = 0 Then
    discount = "0.00"
Else
    discount = dis
End If
End Function

Public Function amt(amount As String) As String
a = InStr(amount, ".")
If Len(amount) = 1 And a = 1 Then
     amt = "0.00"
     Exit Function
End If
If a > 1 Then
    If InStr(a + 1, amount, ".") > 0 Then
        amt = "0.00"
    Else
        amt = Mid(amount, 1, a) & Mid(amount, a + 1, 2)
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
a = InStr(amount, ".")
If Len(amount) = 1 And a = 1 Then
     FAC = "0.000"
     Exit Function
End If
If a > 1 Then
    If InStr(a + 1, amount, ".") > 0 Then
        FAC = "0.000"
    Else
        FAC = Mid(amount, 1, a) & Mid(amount, a + 1, 3)
                    End If
ElseIf a = 1 And Len(amount) = 2 Then
    FAC = "0" & amount & "00"
ElseIf a = 1 And Len(amount) = 3 Then
    FAC = "0" & Mid(amount, 1, 3) & "0"
ElseIf a = 1 And Len(amount) >= 4 Then
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
   MsgBox "Zephyr License Hardware Lock (Dongle) Missing! Check Device Connection to Parallel Port - USB!", vbCritical, "DAS Version 4.0"
   End
End If
Dim dBuf(50) As Byte
dSign = ""
dResult = cl_get_osign(dBuf(0), "")
If dResult = 1 Then
      For I = 0 To 49
        dSign = dSign + Chr(dBuf(I))
      Next
End If
Dim CL_GML As String
Dim mBuf(50) As Byte
dModel = cl_get_model(mBuf(0))
If dModel = 1 Then
For I = 0 To 49
    CL_GML = CL_GML + Chr(mBuf(I))
Next
End If
Dim Dng_Stmp As New ADODB.Recordset
If Trim(Left(dSign, 8)) = "DASHCCBL" Then
    'MsgBox ("DAS Academic Version, HCCBPL Asset"), vbInformation, "DAS Version 4.0"
    M_MDI.Caption = "Zephyr - Integrated Distribution Management" & "Company Asset, Academic License " & "Login : " & Trim(TXTUSERNAME) & " for " & Trim(CompanyName)
    If Dng_Stmp.State = 1 Then Dng_Stmp.Close
    Dng_Stmp.Open "SELECT * FROM DIR  WHERE DIR_IDY = 'DNG' AND DIR_IDX = 'STM'", CON, adOpenKeyset, adLockPessimistic
    If Dng_Stmp.RecordCount > 0 Then
        Dng_Stmp("DIR_ALS") = "Academic"
        Dng_Stmp("DIR_DAT") = Format(Date, "MM/dd/yyyy")
        Dng_Stmp("DIR_MSG") = Trim(CL_GML)
        Dng_Stmp.Update
    Else
        Dng_Stmp.AddNew
        Dng_Stmp("DIR_IDY") = "DNG"
        Dng_Stmp("DIR_IDX") = "STM"
        Dng_Stmp("DIR_ALS") = "Academic"
        Dng_Stmp("DIR_MSG") = Trim(CL_GML)
        Dng_Stmp("DIR_DAT") = Format(Date, "MM/dd/yyyy")
        Dng_Stmp.Update
    End If
Else
    If Dng_Stmp.State = 1 Then Dng_Stmp.Close
    Dng_Stmp.Open "SELECT * FROM DIR  WHERE DIR_IDY = 'DNG' AND DIR_IDX = 'STM'", CON, adOpenKeyset, adLockPessimistic
    If Dng_Stmp.RecordCount > 0 Then
        Dng_Stmp("DIR_ALS") = "Regular"
        Dng_Stmp("DIR_DAT") = Format(Date, "MM/dd/yyyy")
        Dng_Stmp("DIR_MSG") = Trim(CL_GML)
        Dng_Stmp.Update
    Else
        Dng_Stmp.AddNew
        Dng_Stmp("DIR_IDY") = "DNG"
        Dng_Stmp("DIR_IDX") = "STM"
        Dng_Stmp("DIR_ALS") = "Regular"
        Dng_Stmp("DIR_DAT") = Format(Date, "MM/dd/yyyy")
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
   MsgBox "Zephyr License Hardware Lock (Dongle) Missing! Check Device Connection to Parallel Port - USB!", vbCritical, "DAS Version 4.0"
   End
End If
End Sub

Public Sub LOGIN()
If cl_login("067670508880070277000726") <> 1 Then
    MsgBox "Zephyr License Hardware Lock (Dongle) Missing! Check Device Connection to Parallel Port - USB!", vbCritical, "DAS Version 4.0"
End
End If
End Sub

Public Sub LOGOUT()


Dim RSlog As New ADODB.Recordset
RSlog.Open "select * from log", CON, adOpenKeyset, adLockPessimistic
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

Dim rs As New ADODB.Recordset
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY = 'SYS' AND DIR_IDX = 'FTD'") = True Then
    M_ADD_DYS = 0#
    M_ADD_DYS = GetValue("DIR_ALS", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'FTD'")
    If Val(M_ADD_DYS) = 0 Then
        M_ADD_DYS = 0
    Else
        M_ADD_DYS = Val(M_ADD_DYS)
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM YER WHERE ACT_PAS = 'Y'", CON, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then rs.MoveFirst
    For a = 0 To FRM.COUNT - 1
        If TypeOf FRM.Controls(a) Is DTPicker Then
            'FRM.Controls(A).MaxDate = rs("TO_DAT")
            SLD = Format(DateAdd("D", Val(M_ADD_DYS), Date), "MM/dd/yyyy")
            If SLD <= Trim(rs("TO_DAT")) Then
                FRM.Controls(a).MaxDate = SLD
            Else
                FRM.Controls(a).MaxDate = Trim(rs("TO_DAT"))
            End If
            FRM.Controls(a).MinDate = Trim(rs("FRM_DAT"))
            FRM.Controls(a).Value = Date
            CON.Execute "UPDATE DIR SET DIR_ALS = '" & Format(rs("FRM_DAT"), "MM/dd/yyyy") & "' WHERE DIR_IDY = 'CUR'"
            CON.Execute "UPDATE DIR SET DIR_MSG = '" & Format(rs("TO_DAT"), "MM/dd/yyyy") & "' WHERE DIR_IDY = 'CUR'"
        End If
    Next
    Exit Sub
Else
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM YER WHERE ACT_PAS = 'Y'", CON, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then rs.MoveFirst
    For a = 0 To FRM.COUNT - 1
        If TypeOf FRM.Controls(a) Is DTPicker Then
            FRM.Controls(a).MaxDate = Trim(rs("TO_DAT"))
            FRM.Controls(a).MinDate = Trim(rs("FRM_DAT"))
            FRM.Controls(a).Value = Date
            CON.Execute "UPDATE DIR SET DIR_ALS = '" & Format(rs("FRM_DAT"), "MM/dd/yyyy") & "' WHERE DIR_IDY = 'CUR'"
            CON.Execute "UPDATE DIR SET DIR_MSG = '" & Format(rs("TO_DAT"), "MM/dd/yyyy") & "' WHERE DIR_IDY = 'CUR'"
        End If
    Next
    Exit Sub
End If
Label:
MsgBox "Invalid Financial Year & Accounting Period, Contact Zephyr Support!", vbCritical, "DAS Version 4.0"
End
End Sub

Public Sub SETYEARNEW(FRM As Form)
On Error GoTo Label

Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM YER WHERE ACT_PAS = 'Y'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then rs.MoveFirst
For a = 0 To FRM.COUNT - 1
    If TypeOf FRM.Controls(a) Is DTPicker Then
        FRM.Controls(a).MaxDate = rs("FRM_DAT")
        FRM.Controls(a).MinDate = rs("FRM_DAT")
        FRM.Controls(a).Value = rs("FRM_DAT")
        CON.Execute "UPDATE DIR SET DIR_ALS = '" & Format(rs("FRM_DAT"), "MM/dd/yyyy") & "' WHERE DIR_IDY = 'CUR'"
        CON.Execute "UPDATE DIR SET DIR_MSG = '" & Format(rs("TO_DAT"), "MM/dd/yyyy") & "' WHERE DIR_IDY = 'CUR'"
    End If
Next
Exit Sub
Label:
MsgBox "Invalid Accounting Period, Contact Zephyr Support!", vbCritical, "DAS Version 4.0"
End
End Sub
Public Sub SETYEAROPEN(FRM As Form)
On Error GoTo Label

If FRM.Caption = "DB Quick Restore" Then
    Exit Sub
End If
Dim rs As New ADODB.Recordset
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY = 'SYS' AND DIR_IDX = 'AYR'") = True Then
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM YER WHERE ACT_PAS = 'Y'", CON, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then rs.MoveFirst
    For a = 0 To FRM.COUNT - 1
        If TypeOf FRM.Controls(a) Is DTPicker Then
            FRM.Controls(a).MaxDate = Format("31/03/2015", "MM/dd/yyyy")
            FRM.Controls(a).MinDate = Format("01/04/1999", "MM/dd/yyyy")
            FRM.Controls(a).Value = Format(Now, "MM/dd/yyyy")
            CON.Execute "UPDATE DIR SET DIR_ALS = '" & Format(rs("FRM_DAT"), "MM/dd/yyyy") & "' WHERE DIR_IDY = 'CUR'"
            CON.Execute "UPDATE DIR SET DIR_MSG = '" & Format(rs("TO_DAT"), "MM/dd/yyyy") & "' WHERE DIR_IDY = 'CUR'"
        End If
    Next
    Exit Sub
Else
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM YER WHERE ACT_PAS = 'Y'", CON, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then rs.MoveFirst
    For a = 0 To FRM.COUNT - 1
        If TypeOf FRM.Controls(a) Is DTPicker Then
            FRM.Controls(a).MaxDate = Format(rs("TO_DAT"), "MM/dd/yyyy")
            FRM.Controls(a).MinDate = Format(rs("FRM_DAT"), "MM/dd/yyyy")
            FRM.Controls(a).Value = Format(Now, "MM/dd/yyyy")
            CON.Execute "UPDATE DIR SET DIR_ALS = '" & Format(rs("FRM_DAT"), "MM/dd/yyyy") & "' WHERE DIR_IDY = 'CUR'"
            CON.Execute "UPDATE DIR SET DIR_MSG = '" & Format(rs("TO_DAT"), "MM/dd/yyyy") & "' WHERE DIR_IDY = 'CUR'"
        End If
    Next
    Exit Sub
End If
Label:
MsgBox "Invalid Financial Year & Accounting Period, Contact Zephyr Support!", vbCritical, "DAS Version 4.0"
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

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select sub_unt,cur_lev from prd where prd_idy = '" & Trim(PRDID) & "'", CON, adOpenKeyset, adLockPessimistic
Ent_qty = MTS(Val(EQTY), rs(0))
CUR_QTY = MTS(Val(rs(1)), rs(0))
CONQTY = STM(Val(CUR_QTY) - Val(Ent_qty), rs(0))
End Function

Public Function CONPRDQTY(ByVal CURQTY As String, ByVal EQTY As String, PRDID As String) As Long

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
EQTY = QTY(EQTY)
CURQTY = QTY(CURQTY)
rs.Open "select sub_unt from prd where prd_idy = '" & PRDID & "'", CON, adOpenKeyset, adLockPessimistic
If Val(rs(0)) > 0 Then
    a = InStr(EQTY, ".")
    b = InStr(CURQTY, ".")
    If b > 0 Then
        CDECNUM = Mid(CURQTY, b + 1, 3)
        NUM2 = Mid(CURQTY, 1, b - 1)
        TOT1 = Val(rs(0) * NUM2) + Val(CDECNUM)
    End If
    
    If a > 0 Then
    EDECNUM = Mid(EQTY, a + 1, 3)
    NUM1 = Mid(EQTY, 1, a - 1)
    TOT2 = Val(rs(0) * Val(NUM1)) + Val(EDECNUM)
    End If
    CONPRDQTY = TOT1 - TOT2
Else
    CONPRDQTY = QTY(Round(CURQTY - EQTY, 3))
End If
End Function

Public Function ADDQTY(ByVal EQTY As Double, PRDID As String) As Double

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select sub_unt,cur_lev from prd where PRD_idy = '" & Trim(PRDID) & "'", CON, adOpenKeyset, adLockPessimistic
Ent_qty = MTS(Val(EQTY), rs(0))
CUR_QTY = MTS(Val(rs(1)), rs(0))
ADDQTY = STM(Val(Ent_qty) + Val(CUR_QTY), rs(0))
End Function

Public Function ADDBRDQTY(ByVal EQTY As Double, BRDID As String) As Double
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select sub_unt,prd_qty from prd,brd where PRD.PRD_IDY = brd.PRD_idy and BRD.bat_idy = '" & Trim(BRDID) & "'", CON, adOpenKeyset, adLockPessimistic
Ent_qty = MTS(Val(EQTY), rs(0))
CUR_QTY = MTS(Val(rs(1) & ""), rs(0))
ADDBRDQTY = STM(Val(Ent_qty) + Val(CUR_QTY), rs(0))
End Function
', Optional DEFA As String
Public Function CONBRDQTY(ByVal EQTY As Double, BRDID As String) As Double

Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select sub_unt,prd_qty from prd,brd where PRD.PRD_IDY = brd.PRD_idy and bat_idy = '" & Trim(BRDID) & "'", CON, adOpenKeyset, adLockPessimistic
Ent_qty = MTS(Val(EQTY), rs(0))
CUR_QTY = MTS(Val(rs(1)), rs(0))
'If Trim(DEFA) = "+" Then
'    CONBRDQTY = STM(Val(CUR_QTY) + Val(Ent_qty), RS(0))
'ElseIf Trim(DEFA) = "-" Then
    CONBRDQTY = STM(Val(CUR_QTY) - Val(Ent_qty), rs(0))
'End If
End Function

Public Function VANADDQTY(EQTY As String, PRDID As String, invid As String) As String


Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
EQTY = QTY(EQTY)

rs.Open "select sub_unt from prd where prd_idy = '" & Trim(PRDID) & "'", CON, adOpenKeyset, adLockPessimistic
RS1.Open "select RET_QTY from inv where prd_idy = '" & Trim(PRDID) & "' and txn_typ = 'RSI' AND inv_idy = '" & Trim(invid) & "'", CON, adOpenKeyset, adLockPessimistic
If Val(rs(0)) > 0 Then
    a = InStr(QTY(EQTY), ".")
    b = InStr(QTY(RS1(0) & ""), ".")
    
    If a > 0 Then
    EDECNUM = Mid(EQTY, a + 1, 3)
    NUM1 = Mid(EQTY, 1, a - 1)
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


Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
EQTY = QTY(EQTY)

rs.Open "select sub_unt from prd where prd_idy = '" & Trim(PRDID) & "'", CON, adOpenKeyset, adLockPessimistic
RS1.Open "select RET_qty from inv where prd_idy = '" & Trim(PRDID) & "' and txn_typ = 'RSI' AND inv_idy = '" & Trim(invid) & "'", CON, adOpenKeyset, adLockPessimistic
If Val(rs(0)) > 0 Then
    a = InStr(QTY(EQTY), ".")
    b = InStr(QTY(RS1(0) & ""), ".")
    
    If a > 0 Then
    EDECNUM = Mid(EQTY, a + 1, 3)
    NUM1 = Mid(EQTY, 1, a - 1)
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
a = InStr(QTY(unt), ".")
If a > 0 Then
    EDECNUM = Mid(unt, a + 1, 3)
    NUM1 = Mid(unt, 1, a - 1)
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
a = InStr(QTY(unt), ".")
If a > 0 Then
    EDECNUM = Mid(unt, a + 1, 3)
    NUM1 = Mid(unt, 1, a - 1)
End If
WGTCAL = QTY((Val(NUM1) * Val(CONFAC)) + Round(((Val(EDECNUM) * Val(CONFAC)) / Val(SUBUNT)), 3))
End Function

Public Function UNTCAL(unt As String, PRDID As String) As String


Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM PRD WHERE PRD_IDY = '" & Trim(PRDID) & "'", CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
a = Int(Val(unt) / Val(rs("CON_FAC") & ""))
C = Round((Val(rs("CON_FAC") * 1000) / Val(rs("SUB_UNT") * 1000)), 3)
b = Val(unt) * 1000 Mod C * 1000
If b <> 0 Then MsgBox "Conversion Factor and Sub Units Mismatch!", vbExclamation
b = Val(unt) * 1000 Mod Val(rs("CON_FAC") & "") * 1000
b = b / (Val(rs("CON_FAC") & "") / Val(rs("SUB_UNT") & "") * 1000)
b = b / 1000
UNTCAL = QTY(Val(a) + Val(b))
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


Dim MLOS As Long, MPRO As Long
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
Dim RS1 As New ADODB.Recordset
If RS1.State = 1 Then RS1.Close
Dim rs2 As New ADODB.Recordset
If rs2.State = 1 Then rs2.Close
rs.Open "select * from opa", CON, adOpenKeyset, adLockPessimistic
While Not rs.EOF
    If RS1.State = 1 Then RS1.Close
    RS1.Open "select * from act WHERE act_idy = '" & Trim(rs("act_idy") & "") & "'", CON, adOpenKeyset, adLockPessimistic
    If RS1.RecordCount > 0 Then
        If rs2.State = 1 Then rs2.Close
        rs2.Open "select * from typ where typ_def = '" & Trim(RS1("grp_clf") & "") & "'", CON, adOpenKeyset, adLockPessimistic
        If rs2.RecordCount > 0 Then
            If Trim(rs2("typ_fac") & "") = "D" And Trim(rs2("typ_clf") & "") = "PL" Then
                MLOS = MLOS + Val(rs("opa_dbt"))
            ElseIf Trim(rs2("typ_fac") & "") = "C" And Trim(rs2("typ_clf") & "") = "PL" Then
                MPRO = MPRO + Val(rs("opa_crt"))
            End If
        End If
    End If
    rs.MoveNext
Wend
CON.Execute "update dir set dir_als='" & MLOS & "', dir_msg='" & MPRO & "' WHERE DIR_IDY = 'PRO' AND DIR_IDX = 'LOS'"
End Sub

Public Function P_SUB_UNT(QTY As String, SUBQTY) As String
a = InStr(Trim(QTY), ".")
If a > 0 Then
    EDECNUM = Mid(QTY, a + 1, 3)
    NUM1 = Mid(QTY, 1, a - 1)
    P_SUB_UNT = Val(Val(NUM1) * Val(SUBQTY)) + Val(EDECNUM)
End If
End Function

Public Function P_QTY_UNT(M_QTY As String, S_QTY) As String

If Val(M_QTY) < 0 Then
    M_CHK = True
Else
    M_CHK = False
End If
M_QTY = Abs(M_QTY)
If S_QTY <> 1 Then
    M_MOD = Val(M_QTY) Mod S_QTY
    a = (M_QTY - M_MOD)
    b = Val(a) / Val(S_QTY)
    C = Val(b) + (M_MOD / 1000)
Else
    C = M_QTY
End If
If M_CHK = True Then C = C * -1
P_QTY_UNT = QTY(Trim(C))
End Function




Public Sub UPD_STK()
M_FRM_UPD.Show 1

End Sub

Public Sub N_UPD_STK()


If MsgBox("Update Stocks..", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
    Dim REC_SET As New ADODB.Recordset
    Dim REC_SET1 As New ADODB.Recordset
    
    CON.Execute "UPDATE PRD SET CUR_LEV = 0"
    REC_SET.Open "SELECT * FROM INV where txn_typ in('OPS','PUR','SRN','STO','RSR','PCR')", CON, adOpenKeyset, adLockPessimistic
    If REC_SET.EOF = False Then
       While Not REC_SET.EOF
                If REC_SET1.State = 1 Then REC_SET1.Close
                REC_SET1.Open "SELECT * FROM PRD WHERE PRD_IDY = '" & Trim(REC_SET("PRD_IDY")) & "'", CON, adOpenKeyset, adLockPessimistic
                If REC_SET1.RecordCount > 0 Then
                    PRD_SUB = P_SUB_UNT(QTY(REC_SET1("CUR_LEV") & ""), REC_SET1("SUB_UNT"))
                    INV_SUB = P_SUB_UNT(QTY(REC_SET("PRD_QTY") & ""), REC_SET1("SUB_UNT"))
                    REC_SET1("CUR_LEV") = P_QTY_UNT(Val(PRD_SUB) + Val(INV_SUB), REC_SET1("SUB_UNT"))
                    REC_SET1.Update
                End If
                REC_SET.MoveNext
        Wend
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "SELECT * FROM INV where txn_typ in('LSL','SAL','PRT','STI','RSI','PCS')", CON, adOpenKeyset, adLockPessimistic
        While Not REC_SET.EOF
            If REC_SET1.State = 1 Then REC_SET1.Close
            REC_SET1.Open "SELECT * FROM PRD WHERE PRD_IDY = '" & Trim(REC_SET("PRD_IDY")) & "'", CON, adOpenKeyset, adLockPessimistic
            If REC_SET1.RecordCount > 0 Then
                PRD_SUB = P_SUB_UNT(QTY(REC_SET1("CUR_LEV") & ""), REC_SET1("SUB_UNT"))
                INV_SUB = P_SUB_UNT(QTY(REC_SET("PRD_QTY") & ""), REC_SET1("SUB_UNT"))
                REC_SET1("CUR_LEV") = P_QTY_UNT(Val(PRD_SUB) - Val(INV_SUB), REC_SET1("SUB_UNT"))
                REC_SET1.Update
            End If
            REC_SET.MoveNext
        Wend
    End If
End If
End Sub


Public Sub FAupdateCLG(FRM_DAT1 As Date, TO_DATE1 As Date, grpid As String)
FRM_DAT = FRM_DAT1
TO_DATE = TO_DATE1
M_FRM_APD.Tag = "FA_CLG"
M_FRM_APD.Show 1
End Sub
Public Sub FAupdate(FRM_DAT1 As Date, TO_DATE1 As Date, grpid As String, tp As String)
groupID = grpid
ACCOUNTTYP = tp
FRM_DAT = FRM_DAT1
TO_DATE = TO_DATE1
M_FRM_APD.Tag = "FA"
M_FRM_APD.Show 1
End Sub

Public Sub O_UPD_STK()
If MsgBox("Update Stocks..", vbExclamation + vbYesNo, "DAS Version 4.0") = vbYes Then


CON.Execute "UPDATE PRD SET CUR_LEV = 0"
Dim RSI As New ADODB.Recordset
If RSI.State = 1 Then RSI.Close
RSI.Open "SELECT INV.TXN_TYP,INV.PRD_QTY,INV.PRD_IDY,PRD.SUB_UNT FROM INV,PRD WHERE INV.TXN_TYP IN ('OPS','PUR','SRN','STI','RSR') AND PRD.PRD_IDY = INV.PRD_IDY", CON, adOpenKeyset, adLockPessimistic
Do While Not RSI.EOF
    If Val(RSI(3)) <> 1 Then
        CON.Execute "UPDATE PRD SET CUR_LEV= " & ADDQTY(RSI(1), RSI(2)) & "  WHERE PRD_IDY = '" & Trim(RSI(2)) & "'"
    Else
        CON.Execute "UPDATE PRD SET CUR_LEV= CUR_LEV +" & RSI(1) & "  WHERE PRD_IDY = '" & Trim(RSI(2)) & "'"
    End If
    RSI.MoveNext
Loop

If RSI.State = 1 Then RSI.Close
RSI.Open "SELECT INV.TXN_TYP,INV.PRD_QTY,INV.PRD_IDY,PRD.SUB_UNT FROM INV,PRD WHERE INV.TXN_TYP = 'LSL' AND PRD.PRD_IDY = INV.PRD_IDY", CON, adOpenKeyset, adLockPessimistic
Do While Not RSI.EOF
    If Val(RSI(3)) <> 1 Then
        CON.Execute "UPDATE PRD SET CUR_LEV= " & CONQTY(RSI(1), RSI(2)) & "  WHERE PRD_IDY = '" & Trim(RSI(2)) & "'"
    Else
        CON.Execute "UPDATE PRD SET CUR_LEV= CUR_LEV- " & RSI(1) & "  WHERE PRD_IDY = '" & Trim(RSI(2)) & "'"
    End If
    RSI.MoveNext
Loop

If RSI.State = 1 Then RSI.Close
RSI.Open "SELECT INV.TXN_TYP,INV.PRD_QTY,INV.PRD_IDY,PRD.SUB_UNT FROM INV,PRD WHERE INV.TXN_TYP IN ('SAL','RSI','STO','PRT') AND PRD.PRD_IDY = INV.PRD_IDY", CON, adOpenKeyset, adLockPessimistic
Do While Not RSI.EOF
    If Val(RSI(3)) <> 1 Then
        CON.Execute "UPDATE PRD SET CUR_LEV= " & CONQTY(RSI(1), RSI(2)) & "  WHERE PRD_IDY = '" & Trim(RSI(2)) & "'"
    Else
        CON.Execute "UPDATE PRD SET CUR_LEV= CUR_LEV -" & RSI(1) & "  WHERE PRD_IDY = '" & Trim(RSI(2)) & "'"
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
    a = Int(M_QTY) * Val(S_QTY)
    b = (Val(M_QTY) - Int(M_QTY))
    b = b * 1000
    C = Val(a) + Val(b)
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
        a = (M_QTY - M_MOD)
        b = Val(a) / Val(S_QTY)
        C = Val(b) + (M_MOD / 1000)
    Else
        M_MOD = Val(M_QTY) Mod S_QTY
        a = (M_QTY - M_MOD)
        b = Val(a) / Val(S_QTY)
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

Dim rs As New ADODB.Recordset
rs.Open "select " & fldnames & " from " & tblName, CON, adOpenForwardOnly, adLockReadOnly
If rs.EOF = False Then
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

If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
rs.Open "select * from inl", CN1, adOpenKeyset, adLockPessimistic
RS1.Open "select * from pnl", CN1, adOpenKeyset, adLockPessimistic
CN1.Execute "delete from pnl"
While Not rs.EOF
RS1.AddNew
For I = 0 To rs.Fields.COUNT - 1
FLD = rs.Fields(I).name
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
For I = 4 To rs.Fields.COUNT - 1
If rs.Fields(I).Type = adInteger Or rs.Fields(I).Type = adNumeric Or rs.Fields(I).Type = adBigInt Then
CN1.Execute "update pnl set " & rs.Fields(I).name & "=" & rs(I) & "where inv_idy = '" & rs("inv_idy") & "'"
ElseIf rs.Fields(I).Type = adDate Or rs.Fields(I).Type = adDBDate Or rs.Fields(I).Type = adDBTimeStamp Then
CN1.Execute "update pnl set " & rs.Fields(I).name & "='" & Format(rs(I), "MM/dd/yyyy") & "' where inv_idy = '" & rs("inv_idy") & "'"
Else
CN1.Execute "update pnl set " & rs.Fields(I).name & "='" & rs(I) & "' where inv_idy = '" & rs("inv_idy") & "'"
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
RS1(fnm) = ""
Resume Next
ElseIf RS1(fnm).Type = adNumeric Or RS1(fnm).Type = adBigInt Or RS1(fnm).Type = adInteger Or RS1(fnm).Type = adSingle Or RS1(FLD).Type = adDouble Then
RS1(fnm) = 0
Resume Next
ElseIf RS1(fnm).Type = adDate Or RS1(fnm).Type = adDBDate Or RS1(fnm).Type = adDBTimeStamp Then
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

If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
rs.Open "select * from inv", CN1, adOpenKeyset, adLockPessimistic
RS1.Open "select * from pnv", CN1, adOpenKeyset, adLockPessimistic
CN1.Execute "delete from pnv"
While Not rs.EOF
RS1.AddNew
For I = 0 To rs.Fields.COUNT - 1
FLD = rs.Fields(I).name
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
For I = 3 To rs.Fields.COUNT - 1
If rs.Fields(I).Type = adInteger Or rs.Fields(I).Type = adNumeric Or rs.Fields(I).Type = adBigInt Then
CN1.Execute "update pnv set " & rs.Fields(I).name & "=" & rs(I) & "where inv_idy = '" & rs("inv_idy") & "'"
ElseIf rs.Fields(I).Type = adDate Or rs.Fields(I).Type = adDBDate Or rs.Fields(I).Type = adDBTimeStamp Then
CN1.Execute "update pnv set " & rs.Fields(I).name & "='" & Format(rs(I), "MM/dd/yyyy") & "' where inv_idy = '" & rs("inv_idy") & "'"
Else
CN1.Execute "update pnv set " & rs.Fields(I).name & "='" & rs(I) & "' where inv_idy = '" & rs("inv_idy") & "'"
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
RS1(fnm) = ""
Resume Next
ElseIf RS1(fnm).Type = adNumeric Or RS1(fnm).Type = adBigInt Or RS1(fnm).Type = adInteger Or RS1(fnm).Type = adSingle Or RS1(FLD).Type = adDouble Then
RS1(fnm) = 0
Resume Next
ElseIf RS1(fnm).Type = adDate Or RS1(fnm).Type = adDBDate Or RS1(fnm).Type = adDBTimeStamp Then
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


If rs2.State = 1 Then rs2.Close
    If rs3.State = 1 Then rs3.Close
    rs3.Open "SELECT * FROM PRD", CN1, adOpenDynamic, adLockPessimistic
    While Not rs3.EOF
        CN1.Execute "update inv set sub_unt = " & Val(rs3!SUB_UNT) & " where prd_idy = '" & Trim(rs3!prd_idy) & "'", d
        CN1.Execute "update inv set prd_uom = " & Val(rs3!con_fac) & " where prd_idy = '" & Trim(rs3!prd_idy) & "'", d
        rs3.MoveNext
    Wend
    CN1.Execute "update inv set stk_typ='Salable'"
    CN1.Execute "update inl set txn_typ = 'SAL' WHERE TXN_TYP = 'LSL'"
    CN1.Execute "update inv set txn_typ = 'SAL' WHERE TXN_TYP = 'LSL'"
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
rs.Open "select * from rdy where act_idy = '" & Trim(CustomerID) & "' and smn_idy = '" & Trim(SalesmanID) & "' and day_idy =" & d, CON, adOpenKeyset, adLockPessimistic
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
rs.Open "select * from rdy,smn where smn.smn_idy = rdy.smn_idy and rdy.act_idy = '" & Trim(CustomerID) & "' and day_idy =" & d, CON, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then GetSalesman = Trim(rs!SMN_IDY & "") & "|" & Trim(rs!SMN_NME & "")
End Function

Public Sub AddShells(Invoice As String, Txntype As String, PartyID As String, UPD As String, M_SMN_IDY As String, M_ROU_IDY As String)
Dim RSINV As New ADODB.Recordset
Dim RSSNV As New ADODB.Recordset
CON.Execute "DELETE FROM SNV WHERE INV_IDY = '" & Trim(Invoice) & "' AND TXN_TYP = '" & Trim(Txntype) & "'"
If Txntype = "SAL" Then
    RSINV.Open "SELECT sum(tot_pcs),inv.prd_idy,inv.inv_idy,inv.inv_dat,shl.shl_idy,prd.sub_unt,prd.prd_pck,INV.ORD_IDY FROM INV,SHL,PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND INV.PRD_IDY = SHL.PRD_IDY AND INV_IDY = '" & Trim(Invoice) & "' AND inv.TXN_TYP = '" & Trim(Txntype) & "' AND SHL_TYP ='Shell' AND SHL_TYP  IN('Salable','Free') group by inv.prd_idy,inv.inv_idy,inv.inv_dat,shl.shl_idy,prd.sub_unt,prd.prd_pck,INV.ORD_IDY", CON, adOpenKeyset, adLockPessimistic
Else
    RSINV.Open "SELECT sum(tot_pcs),inv.prd_idy,INV.inv_idy,INV.inv_dat,SHL.shl_idy,PRD.sub_unt,prd.PRD_Pck,INV.ORD_IDY FROM INV,SHL,PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND INV.PRD_IDY = SHL.PRD_IDY AND INV_IDY = '" & Trim(Invoice) & "' AND INV.TXN_TYP = '" & Trim(Txntype) & "' AND SHL_TYP ='Shell' AND SHL_TYP  NOT IN('Breakage') group by inv.prd_idy,INV.inv_idy,INV.inv_dat,SHL.shl_idy,PRD.sub_unt,prd.PRD_Pck,INV.ORD_IDY", CON, adOpenKeyset, adLockPessimistic
End If
RSSNV.Open "SELECT * FROM SNV ", CON, adOpenKeyset, adLockPessimistic
For a = 1 To RSINV.RecordCount
    If Txntype = "SAL" Then If a = 1 Then PRDID = RSINV!prd_idy         ' TO DISCUSS WITH REDDY SIR
    RSSNV.AddNew
    RSSNV!SMN_IDY = Trim(M_SMN_IDY)
    RSSNV!rou_idy = Trim(M_ROU_IDY)
    RSSNV!INV_IDY = Trim(RSINV!INV_IDY)
    RSSNV!INV_DAT = Trim(RSINV!INV_DAT)
    RSSNV!prd_idy = Trim(RSINV!prd_idy)
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
    RSINV.Open "SELECT sum(tot_pcs),inv.prd_idy,inv.inv_idy,inv.inv_dat,shl.shl_idy,prd.sub_unt,prd.prd_pck,INV.ORD_IDY FROM INV,SHL,PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND INV.PRD_IDY = SHL.PRD_IDY AND INV_IDY = '" & Trim(Invoice) & "' AND inv.TXN_TYP = '" & Trim(Txntype) & "' AND SHL_TYP ='Shell' AND SHL_TYP  IN('Damage','TakeBack','Breakage','Leakage') group by inv.prd_idy,inv.inv_idy,inv.inv_dat,shl.shl_idy,prd.sub_unt,prd.prd_pck,INV.ORD_IDY ", CON, adOpenKeyset, adLockPessimistic
    For a = 1 To RSINV.RecordCount
        CON.Execute "update snv set i_shl_nos = i_shl_nos - " & Int(Val(STM(Val(RSINV(0) & ""), Val(RSINV(5))))) & " where inv_idy = '" & Trim(RSINV!INV_IDY) & "' and prd_idy  ='" & RSINV!prd_idy & "'", d
        CON.Execute "update snv set i_pck_nos = i_pck_nos - " & Val(RSINV(0) & "") & " where inv_idy = '" & Trim(RSINV!INV_IDY) & "' and prd_idy  ='" & RSINV!prd_idy & "'", e
        CON.Execute "update snv set r_shl_nos = r_shl_nos - " & Int(Val(STM(Val(RSINV(0) & ""), Val(RSINV(5))))) & " where inv_idy = '" & Trim(RSINV!INV_IDY) & "' and prd_idy  ='" & RSINV!prd_idy & "'", f
        CON.Execute "update snv set r_pck_nos = r_pck_nos - " & Val(RSINV(0) & "") & " where inv_idy = '" & Trim(RSINV!INV_IDY) & "' and prd_idy  ='" & RSINV!prd_idy & "'", g
        If d = 0 Or e = 0 Or f = 0 Or g = 0 Then
            RSSNV.AddNew
            RSSNV!SMN_IDY = Trim(M_SMN_IDY)
            RSSNV!rou_idy = Trim(M_ROU_IDY)
            RSSNV!INV_IDY = Trim(RSINV!INV_IDY)
            RSSNV!INV_DAT = Trim(RSINV!INV_DAT)
            RSSNV!prd_idy = Trim(RSINV!prd_idy)
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
RSINV.Open "SELECT sum(tot_pcs),inv.prd_idy,INV.inv_idy,INV.inv_dat,SHL.shl_idy,PRD.sub_unt,prd.prd_pck FROM INV,SHL,prd WHERE inv.prd_idy = PRD.PRD_IDY and INV.PRD_IDY = SHL.PRD_IDY AND INV_IDY = '" & Trim(Invoice) & "' AND TXN_TYP = '" & Trim(Txntype) & "' AND SHL_TYP ='Shell' AND SHL_TYP  IN ('Breakage') group by inv.prd_idy,INV.inv_idy,INV.inv_dat,SHL.shl_idy,PRD.sub_unt,prd.prd_pck", CON, adOpenKeyset, adLockPessimistic
For a = 1 To RSINV.RecordCount
    CON.Execute "update snv set r_shl_nos=r_shl_nos - " & Int(STM(RSINV(0), RSINV(5))) & ",r_pck_nos=r_pck_nos-" & Val(RSINV(0)) & ",brk_shl=" & Int(STM(RSINV(0), RSINV(5))) & ",brk_pck=" & Val(RSINV(0)) & "where prd_idy = '" & Trim(RSINV(1)) & "' and inv_idy = '" & Trim(Invoice) & "' and txn_typ = '" & Trim(Txntype) & "'", d
    'CON.Execute "update snv set r_shl_nos=i_shl_nos,r_pck_nos=i_pck_nos where prd_idy = '" & Trim(RSINV(1)) & "' and inv_idy = '" & Trim(Invoice) & "' and txn_typ = '" & Trim(Txntype) & "'"
    If d = 0 Then
        CON.Execute "update snv set r_shl_nos=r_shl_nos-" & Int(STM(RSINV(0), RSINV(5))) & ",r_pck_nos=r_pck_nos-" & Val(RSINV(0)) & " where prd_idy = '" & Trim(PRDID) & "' and inv_idy = '" & Trim(Invoice) & "' and txn_typ = '" & Trim(Txntype) & "'", d
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
        RSSNV!prd_idy = Trim(RSINV!prd_idy)
        RSSNV!txn_dat = Trim(RSINV!INV_DAT)
        RSSNV!PRD_PCK = Trim(RSINV!PRD_PCK)
        RSSNV!prd_upc = Val(RSINV!SUB_UNT)
        RSSNV!stk_upd = UPD
        RSSNV.Update
    End If
    RSINV.MoveNext
Next
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
RST_SMN.Open "SELECT * FROM SMN", CON, adOpenKeyset, adLockPessimistic
If RST_SMN.RecordCount > 0 Then RST_SMN.MoveFirst
While Not RST_SMN.EOF
    If RST_RDY.State = 1 Then RST_RDY.Close
    RST_RDY.Open "SELECT  DISTINCT COUNT(RDY.ACT_IDY), DAY_IDY FROM RDY,ACT WHERE RDY.SMN_IDY = '" & RST_SMN("SMN_IDY") & "' AND RDY.ACT_IDY = ACT.ACT_IDY AND ACT.DBT_TYP = 'Y' GROUP BY RDY.SMN_IDY, DAY_IDY", CON, adOpenKeyset, adLockPessimistic
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
Dim RSrP1 As New ADODB.Recordset
Dim DT1 As Date
Dim DT2 As Date
Dim DAYCOUNT(6) As Integer
CON_APP = SMNorROUorACTorDAY
DT1 = FrmDate
DT2 = ToDate
CON.Execute "DELETE FROM RP1"
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

If RSrP1.State = 1 Then RSrP1.Close
RSrP1.Open "SELECT * FROM RP1", CON, adOpenKeyset, adLockPessimistic

If CON_APP = "SMN" Then
    If RSSMN.State = 1 Then RSSMN.Close
    RSSMN.Open "SELECT * FROM SMN", CON, adOpenKeyset
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT COUNT(*),SMN_IDY,DAY_IDY FROM RDY GROUP BY DAY_IDY,SMN_IDY", CON, adOpenKeyset, adLockPessimistic
    While Not RSSMN.EOF
        CTR = 0
        For dayid = 0 To 6
            rs.filter = "SMN_IDY = '" & Trim(RSSMN!SMN_IDY) & "' AND DAY_IDY =" & dayid + 1
            If rs.RecordCount > 0 Then CTR = CTR + (Val(rs(0)) * DAYCOUNT(dayid))
        Next
        RSrP1.AddNew
        RSrP1(0) = ""
        RSrP1(1) = ""
        RSrP1(2) = ""
        RSrP1(3) = ""
        RSrP1(4) = ""
        RSrP1(5) = ""
        RSrP1(6) = ""
        RSrP1(7) = ""
        RSrP1(8) = ""
        RSrP1(9) = RSSMN!SMN_IDY
        RSrP1(10) = 0
        RSrP1(11) = CTR
        RSrP1(12) = 0
        RSrP1(13) = 0
        RSrP1(14) = 0
        RSrP1(15) = 0
        RSrP1(16) = 0
        RSrP1(17) = 0
        RSrP1(18) = 0
        RSrP1(19) = 0
        RSrP1.Update
        RSSMN.MoveNext
    Wend
ElseIf CON_APP = "ROU" Then
    If RSSMN.State = 1 Then RSSMN.Close
    RSSMN.Open "SELECT ROU_IDY FROM RDY GROUP BY ROU_IDY", CON, adOpenKeyset
    For a = 1 To RSSMN.RecordCount
        CTR = 0
        For dayid = 0 To 6
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT COUNT(*),ROU_IDY,DAY_IDY FROM RDY where rou_idy='" & Trim(RSSMN!rou_idy) & "' and day_idy=" & dayid + 1 & " GROUP BY DAY_IDY,ROU_IDY ", CON, adOpenKeyset, adLockPessimistic
            'RS.filter = "ROU_IDY = '" & Trim(RSSMN!ROU_IDY) & "' AND DAY_IDY =" & dayid + 1
            If rs.RecordCount > 0 Then CTR = CTR + (Val(rs(0)) * DAYCOUNT(dayid))
        Next
        RSrP1.AddNew
        RSrP1(0) = ""
        RSrP1(1) = ""
        RSrP1(2) = ""
        RSrP1(3) = ""
        RSrP1(4) = ""
        RSrP1(5) = ""
        RSrP1(6) = ""
        RSrP1(7) = ""
        RSrP1(8) = ""
        RSrP1(9) = Trim(RSSMN!rou_idy & "")
        RSrP1(10) = 0
        RSrP1(11) = CTR
        RSrP1(12) = 0
        RSrP1(13) = 0
        RSrP1(14) = 0
        RSrP1(15) = 0
        RSrP1(16) = 0
        RSrP1(17) = 0
        RSrP1(18) = 0
        RSrP1(19) = 0
        RSrP1.Update
        RSSMN.MoveNext
    Next
ElseIf CON_APP = "ACT" Then
    If RSSMN.State = 1 Then RSSMN.Close
    RSSMN.Open "SELECT ACT_IDY FROM ACT WHERE ACT_TYP IN('BOT','SAL') and cst_typ not in('Van') AND DBT_TYP = 'Y'", CON, adOpenKeyset
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT COUNT(*),ACT_IDY,DAY_IDY FROM RDY GROUP BY DAY_IDY,ACT_IDY", CON, adOpenKeyset, adLockPessimistic
    For a = 1 To RSSMN.RecordCount
        CTR = 0
        For dayid = 0 To 6
            rs.filter = "ACT_IDY = '" & Trim(RSSMN!ACT_IDY) & "' AND DAY_IDY =" & dayid + 1
            If rs.RecordCount > 0 Then CTR = CTR + (Val(rs(0)) * DAYCOUNT(dayid))
        Next
        RSrP1.AddNew
        RSrP1(0) = ""
        RSrP1(1) = ""
        RSrP1(2) = ""
        RSrP1(3) = ""
        RSrP1(4) = ""
        RSrP1(5) = ""
        RSrP1(6) = ""
        RSrP1(7) = ""
        RSrP1(8) = ""
        RSrP1(9) = RSSMN!ACT_IDY
        RSrP1(10) = 0
        RSrP1(11) = CTR
        RSrP1(12) = 0
        RSrP1(13) = 0
        RSrP1(14) = 0
        RSrP1(15) = 0
        RSrP1(16) = 0
        RSrP1(17) = 0
        RSrP1(18) = 0
        RSrP1(19) = 0
        RSrP1.Update
        RSSMN.MoveNext
    Next
ElseIf CON_APP = "DAY" Then
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT COUNT(*),DAY_IDY FROM RDY GROUP BY DAY_IDY", CON, adOpenKeyset, adLockPessimistic
    For a = 1 To RSSMN.RecordCount
        RSrP1.AddNew
        RSrP1(0) = ""
        RSrP1(1) = ""
        RSrP1(2) = ""
        RSrP1(3) = ""
        RSrP1(4) = ""
        RSrP1(5) = ""
        RSrP1(6) = ""
        RSrP1(7) = ""
        RSrP1(8) = ""
        RSrP1(9) = Trim(rs!DAY_IDY & "")
        RSrP1(10) = 0
        RSrP1(11) = Val(rs(0) & "")
        RSrP1(12) = 0
        RSrP1(13) = 0
        RSrP1(14) = 0
        RSrP1(15) = 0
        RSrP1(16) = 0
        RSrP1(17) = 0
        RSrP1(18) = 0
        RSrP1(19) = 0
        RSrP1.Update
        RSSMN.MoveNext
    Next
End If
End Sub

Public Sub doPent(FrmDate As Date, ToDate As Date, SMNorROU As String)
st = SMNorROU
Dim rs As New ADODB.Recordset
Dim RP1 As New ADODB.Recordset
STR1 = GetValue("DIR_MSG", "DIR WHERE DIR_IDY = 'BRD' AND DIR_IDX = 'IDY'")
STR2 = Split(STR1, ",")
If RP1.State = 1 Then RP1.Close
RP1.Open "select * from rp1", CON, adOpenKeyset, adLockPessimistic
For a = 1 To RP1.RecordCount
    If st = "SMN" Then
        qry = " AND  SMN_IDY = '" & Trim(RP1!CHR_TEN) & "'"
    ElseIf st = "ROU" Then
        qry = " AND  ROU_IDY = '" & Trim(RP1!CHR_TEN) & "'"
    End If
    For b = 0 To UBound(STR2)
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT COUNT(DISTINCT(CST_IDY)) FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND PRD.BRD_NME = '" & Trim(STR2(b)) & "' AND INV_DAT >= '" & Format(FrmDate, "MM/dd/yyyy") & "' AND INV_DAT<='" & Format(ToDate, "MM/dd/yyyy") & "' )" & qry, CON, adOpenKeyset, adLockPessimistic
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
RP1.Open "select * from rp1", CON, adOpenKeyset, adLockPessimistic
While Not RP1.EOF
    If st = "SMN" Then
        qry = " AND  SMN_IDY = '" & Trim(RP1!CHR_TEN) & "'"
    ElseIf st = "ROU" Then
        qry = " AND  ROU_IDY = '" & Trim(RP1!CHR_TEN) & "'"
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT COUNT(DISTINCT(ACT_IDY)) FROM RDY WHERE ACT_IDY NOT IN(SELECT CST_IDY FROM INL WHERE INV_DAT >='" & Format(FrmDate, modGeneral.OraDtFormat) & "' AND INV_DAT <='" & Format(ToDate, modGeneral.OraDtFormat) & "'  " & qry & ")" & qry, CON, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        RP1("NUM_THR") = Val(rs(0) & "")
        RP1.Update
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT COUNT(DISTINCT(CST_IDY)) FROM INL WHERE INV_DAT >='" & Format(FrmDate, modGeneral.OraDtFormat) & "' AND INV_DAT <='" & Format(ToDate, modGeneral.OraDtFormat) & "'  " & qry, CON, adOpenKeyset, adLockPessimistic
    If rs.RecordCount > 0 Then
        RP1("NUM_FOR") = Val(rs(0) & "")
        RP1.Update
      End If
    RP1.MoveNext
Wend
End Sub
Public Sub doPentSKU(FrmDate As Date, ToDate As Date, SMNorROU As String)
st = SMNorROU
Dim rs As New ADODB.Recordset
Dim RP1 As New ADODB.Recordset
STR1 = GetValue("DIR_MSG", "DIR WHERE DIR_IDY = 'PRD' AND DIR_IDX = 'IDY'")
STR2 = Split(STR1, ",")
If RP1.State = 1 Then RP1.Close
RP1.Open "select * from rp1", CON, adOpenKeyset, adLockPessimistic
For a = 1 To RP1.RecordCount
    If st = "SMN" Then
        qry = " AND  SMN_IDY = '" & Trim(RP1!CHR_TEN) & "'"
    ElseIf st = "ROU" Then
        qry = " AND  ROU_IDY = '" & Trim(RP1!CHR_TEN) & "'"
    End If
    For b = 0 To UBound(STR2)
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT COUNT(DISTINCT(CST_IDY)) FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND PRD.PRD_IDY = '" & Trim(STR2(b)) & "' AND INV_DAT >= '" & Format(FrmDate, "MM/dd/yyyy") & "' AND INV_DAT<='" & Format(ToDate, "MM/dd/yyyy") & "' )" & qry, CON, adOpenKeyset, adLockPessimistic
        If rs.RecordCount > 0 Then
            RP1(12 + b) = Val(rs(0) & "")
            RP1.Update
        End If
    Next
    RP1.MoveNext
Next
End Sub

Public Function AMT1(amount As String) As String
a = InStr(amount, ".")
If Len(amount) = 1 And a = 1 Then
     AMT1 = "0.00"
     Exit Function
End If
If a > 1 Then
    If InStr(a + 1, amount, ".") > 0 Then
        AMT1 = "0.00"
    Else
        AMT1 = Mid(amount, 1, a) & Mid(amount, a + 1, 2)
    End If
ElseIf a = 1 And Len(amount) = 2 Then
    AMT1 = "0" & amount & "0"
ElseIf a = 1 And Len(amount) >= 3 Then
    AMT1 = "0" & Mid(amount, 1, 3)
ElseIf a = 0 And Len(Trim(amount)) >= 7 Then
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
End Sub
Public Sub TxnMoveBackup(SRC_DB As String, TGT_DB As String, B_TXN_IDY As String, B_TXN_TYP As String)

'Code by PNR for DataBackup - Deleted Recordset before Deletion for Critical Tables in Schema
'Restore Manager Can Restore from these tables upon Sync Verification
Dim SRC_RSET As New ADODB.Recordset
Dim TGT_RSET As New ADODB.Recordset


If SRC_RSET.State = 1 Then SRC_RSET.Close
'If Trim(B_TXN_TYP) = "" Then
    If Trim(SRC_DB) = "LNV" Or Trim(SRC_DB) = "CNV" Then
        SRC_RSET.Open "SELECT  DISTINCT * FROM " & Trim(SRC_DB) & " where INV_IDY = '" & Trim(B_TXN_IDY) & "' ORDER BY LEV_DES", CON, adOpenForwardOnly, adLockReadOnly
    End If
    If Trim(SRC_DB) = "VAN" Or Trim(SRC_DB) = "VBK" Then
        SRC_RSET.Open "SELECT  DISTINCT * FROM " & Trim(SRC_DB) & " where INV_IDY = '" & Trim(B_TXN_IDY) & "' ORDER BY BAT_IDY", CON, adOpenForwardOnly, adLockReadOnly
    End If
    If Trim(SRC_DB) = "SNV" Or Trim(SRC_DB) = "SBK" Then
        If SRC_RSET.State = 1 Then SRC_RSET.Close
        SRC_RSET.Open "SELECT  DISTINCT * FROM " & Trim(SRC_DB) & " where TXN_IDY = '" & Trim(B_TXN_IDY) & "' ORDER BY PRD_PCK", CON, adOpenForwardOnly, adLockReadOnly
    End If
'Else
    If Trim(SRC_DB) = "INV" Or Trim(SRC_DB) = "PNV" Or Trim(SRC_DB) = "XNV" Then
        SRC_RSET.Open "SELECT  DISTINCT * FROM " & Trim(SRC_DB) & " where INV_IDY = '" & Trim(B_TXN_IDY) & "' AND TXN_TYP = '" & Trim(B_TXN_TYP) & "' ORDER BY BAT_IDY", CON, adOpenForwardOnly, adLockReadOnly
    End If
    If Trim(SRC_DB) = "INL" Or Trim(SRC_DB) = "PNL" Or Trim(SRC_DB) = "XNL" Then
        SRC_RSET.Open "SELECT  DISTINCT * FROM " & Trim(SRC_DB) & " where INV_IDY = '" & Trim(B_TXN_IDY) & "' AND TXN_TYP = '" & Trim(B_TXN_TYP) & "' ORDER BY INV_IDY", CON, adOpenForwardOnly, adLockReadOnly
    End If
'End If
If SRC_RSET.State = 1 Then
If TGT_RSET.State = 1 Then TGT_RSET.Close
TGT_RSET.Open "SELECT * FROM " & Trim(TGT_DB) & " ", CON, adOpenDynamic, adLockOptimistic
While Not SRC_RSET.EOF
    TGT_RSET.AddNew
    For K = 0 To TGT_RSET.Fields.COUNT - 1
        If TGT_RSET(K).Type = adVarChar Or TGT_RSET(K).Type = adBSTR Or TGT_RSET(K).Type = adChar Or TGT_RSET(K).Type = adLongVarChar Or TGT_RSET(K).Type = adSmallInt Then
            TGT_RSET(K) = ""
        ElseIf TGT_RSET(K).Type = adInteger Or TGT_RSET(K).Type = adNumeric Or TGT_RSET(K).Type = adBigInt Or TGT_RSET(K).Type = adSingle Or TGT_RSET(K).Type = adDouble Then
            TGT_RSET(K) = 0
        ElseIf TGT_RSET(K).Type = adDate Or TGT_RSET(K).Type = adDBDate Or TGT_RSET(K).Type = adDBTimeStamp Then
            TGT_RSET(K) = Date
        End If
    Next
    For K = 0 To SRC_RSET.Fields.COUNT - 1
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


Public Sub clrctr(FRM As Form)
For a = 1 To FRM.COUNT - 1
If TypeOf FRM.Controls(a) Is TextBox Then FRM.Controls(a).Text = ""
If TypeOf FRM.Controls(a) Is MaskEdBox Then FRM.Controls(a).Text = ""
Next
End Sub

Public Sub SCl(M As MSFlexGrid, C As Integer, Optional r = 0)
M.Row = r
M.Col = C
M.CellForeColor = vbBlue
End Sub

Public Sub M_HELP()
On Error GoTo LBL
M_PATH = RootPath & "\CONTHELP.CHM"
If M_MDI.SysInfo1.OSPlatform = 1 Then
    Shell "C:\WINDOWS\HH.EXE " & M_PATH, vbMaximizedFocus
ElseIf M_MDI.SysInfo1.OSPlatform = 2 Then
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
M_PATH = RootPath & "\CONTHELP.CHM"
If M_MDI.SysInfo1.OSPlatform = 1 Then
    Shell "C:\WINDOWS\HH.EXE " & M_PATH, vbMaximizedFocus
ElseIf M_MDI.SysInfo1.OSPlatform = 2 Then
    Shell "C:\WINNT\HH.EXE " & M_PATH, vbMaximizedFocus
Else
    Shell "C:\WINDOWS\HH.EXE " & M_PATH, vbMaximizedFocus
End If
Exit Sub
LBL:
'MsgBox err.Description, vbCritical
MsgBox "Help will be made available shortly!", vbInformation, "DAS Version 4.0"
End Sub

Public Function CHECKTB(sFile)
tbN = Split(sFile, ",", , vbTextCompare)
For b = 0 To UBound(tbN)
    strPath = CurDir & "\" & Trim(tbN(b))
    If Dir(strPath) = "" Then
        CHECKTB = 1
        TNAME = tbN(b)
        Exit Function
    End If
Next
End Function

Public Function Openreport(ReportLibrary As String, reportname As String, query As Integer, filter As String, destination As Integer)
With M_MDI.RRReport1
    .ReportLibrary = Trim(ReportLibrary)
    .reportname = Trim(reportname)
    .printer = "?"
    .query = Val(query)
    .filter = filter
    .destination = Val(destination)
    .RunReport 1
End With
End Function
