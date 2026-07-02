Attribute VB_Name = "modGeneral"
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Const LOCALE_SSHORTDATE = &H1F

Public OraConnection As ADODB.Connection

Dim IPAddress As String
Dim Provider As String
Dim DataSource As String

Public OraServerConnection As ADODB.Connection
Public OraclePath As String
Public ConnectionString As String
Public DexFolder As String
Public ReportFolder As String
Public SetupFolder As String
Public RootPath As String
Public BackupFolder As String
Public AutoBackupFolder As String
Public AutoRestoreFolder As String
Public HHTFolder As String
Public TempFolder As String
Public InBoxFolder As String
Public OutBoxFolder As String
Public DBFolder As String
Public ServerName As String
Public LogFolder As String
Public SqlplusFolder As String
Public DBUserName As String
Public DBPassword As String
Public DriverName As String
Public DatabaseName As String
Public TNSName As String
Public SMSFolder As String
Public PrinterPath As String
Public DtFormat As String
Public PrintMode As String
Public OraDtFormat As String
Public WindowsServer As String
Public strMsg As String

Public ReadCursorType As Long
Public WriteCursorType As Long

Dim FSO As New FileSystemObject
Dim f As File
Dim ts As TextStream




Public Function SetDirectoryPath() As Long
On Error GoTo ErrHandler
Dim lngLocale As Long


    
    Dim FSO As New FileSystemObject
    Dim ts As TextStream
    RootPath = App.Path
    SetupFolder = RootPath & "\Setup"
    Set ts = FSO.OpenTextFile(SetupFolder & "\Path.dat", ForReading)
    HHTFolder = ts.ReadLine()
    LogFolder = ts.ReadLine()
    SMSFolder = ts.ReadLine()
    If ts.AtEndOfLine = False Then strMsg = ts.ReadLine()
    ts.Close
    
    HHTFolder = App.Path & "\HHT"
    LogFolder = App.Path & "\LOG"
    SMSFolder = App.Path & "\SMS"
    

    lngLocale = GetSystemDefaultLCID()
    SetLocaleInfo lngLocale, LOCALE_SSHORTDATE, "MM/dd/yyyy"


    If Dir(HHTFolder, vbDirectory) = "" Then MkDir HHTFolder
    If Dir(LogFolder, vbDirectory) = "" Then MkDir LogFolder
    If Dir(SMSFolder, vbDirectory) = "" Then MkDir SMSFolder
    
    
    SetDirectoryPath = 0
    Exit Function
ErrHandler:
    SetDirectoryPath = err.Number
    AddPSSErrorLog "ModGeneral - GetRecords"
End Function

Public Function CompanyList(cmb As ComboBox)
    cmb.CLEAR
''    SetConnectionString EFF_LIB1.CompanyName
    CreateSystemDSN "EXISTING COMPANY"
    Dim rec As New ADODB.Recordset
    If rec.State = 1 Then rec.Close
    rec.Open "SELECT  * FROM CML", OraServerConnection, adOpenKeyset, adLockOptimistic
    cmb.AddItem "EXISTING COMPANY"
    While Not rec.EOF
        cmb.AddItem Trim(rec!CMP_NME)
        CreateSystemDSN rec!CMP_NME
        rec.MoveNext
    Wend
    rec.Close
End Function

Public Function GetDBPath()
    
    SetDirectoryPath

End Function

Public Function SetConnectionString(CompanyName As String)
    ConnectionString = "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & EFF_LIB1.CompanyFolderPath & "\EFF.DBC;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
End Function

Public Function SetConnectionString1(CompanyName As String) As String
    Dim FSO As New FileSystemObject
    Dim ts As TextStream
    Dim STR As String
    CompanyName = UCase(RemoveSpaces(CompanyName))
    SetDirectoryPath
    Set ts = FSO.OpenTextFile(SetupFolder & "/DB.dat", ForReading)
    DataSource = ts.ReadLine()
    Provider = ts.ReadLine()
    ServerName = ts.ReadLine()
    ts.Close
    DataSource = Replace(DataSource, "DataSource:", "")
    Provider = Replace(Provider, "Provider:", "")
    If Provider <> "" Then Provider = Provider & ";"
    ServerName = Replace(ServerName, "Server:", "")
    If ServerName <> "" Then ServerName = ServerName & ";"
    STR = EFF_LIB1.CompanyName
    STR = Replace(STR, " ", "")
    SetConnectionString1 = Provider & DataSource & ";" & " User ID=USRPVM" & Trim(CompanyName) & ";" & " Password=PWDPVM" & Trim(CompanyName) & " ; Persist Security Info=TRUE"
End Function



Public Function OraConnect()
    If CompanyName = "" Then Exit Function
    Set OraConnection = New ADODB.Connection
    OraConnection.ConnectionString = ConnectionString
    If OraConnection.State = 1 Then OraConnection.Close
    OraConnection.Open
End Function

Public Function ExecuteSql(SqlStatement As String, Optional bClose As Boolean) As Long
On Error GoTo ErrHandler
    If OraConnection.State = 0 Then OraConnection.Open
    OraConnection.Execute SqlStatement
    If bClose = True Then OraConnection.Close
    ExecuteSql = 0
    Exit Function
ErrHandler:
    ExecuteSql = err.Number
    AddPSSErrorLog "ModGeneral - ExecuteSQL"
End Function

Public Function AddPSSErrorLog(TXT As String)
    Set ts = FSO.OpenTextFile(SetupFolder & "/PSSErrorLog.txt" & fileName, ForAppending, True)
    ts.WriteLine (TXT)
    ts.Close
End Function

Public Function GetRecords(SqlStatement As String, Optional bClose As Boolean) As ADODB.Recordset
On Error GoTo ErrHandler
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    If OraConnection = Null Then OraConnect
    If OraConnection.State = 0 Then OraConnection.Open ConnectionString
'    OraConnection.Open ConnectionString
    rec.CursorLocation = adUseClient
    SqlStatement = Replace(SqlStatement, """" & """", """")
    rec.Open SqlStatement, OraConnection, adOpenForwardOnly, adLockPessimistic
    Set GetRecords = rec
    Exit Function
ErrHandler:
    Set GetRecords = Nothing
    AddPSSErrorLog "ModGeneral - GetRecords"
End Function


Public Function RemoveSpaces(Value As String) As String
    Value = Replace(Value, " ", "")
    RemoveSpaces = Value
End Function

Public Function FoxToOra(CompanyName As String) As String
    Dim UserName As String
    Dim Password As String
    Dim FSO As New FileSystemObject
    Dim ts As TextStream
    UserName = "USRPVM" & UCase(CompanyName)
    Password = "PWDPVM" & UCase(CompanyName)
    Dim STR As String
    STR = STR & " sqlplus " & UserName & "/" & Password & IIf(ServerName <> "", "@" & ServerName & "  ", "  ") & "@" & DBFolder & "/a.sql " & vbCrLf
    Set ts = FSO.CreateTextFile(DBFolder & "/sqlExe.cmd", True)
    ts.Write STR
    ts.Close
    Shell DBFolder & "/sqlExe.cmd", vbMaximizedFocus
End Function

Public Function ReSetConnectionObjects(CompanyName As String)
    SetConnectionString CompanyName
    OraConnect
    Set CON = New ADODB.Connection

    
    CON.Open modGeneral.ConnectionString

End Function


Public Function CloseConnectionObjects()
    SetConnectionString EFF_LIB1.CompanyName
    OraConnection.Close
    CON.Close
    Set CON = Nothing
    Set OraConnection = Nothing
End Function


Public Function ProperCase(STR As String) As String
STR = UCase(Mid(STR, 1, 1)) & Mid(STR, 2)
STR = Replace(STR, "_", " ")
For I = 2 To Len(STR)
    If Asc(Mid(STR, I, 1)) = vbKeySpace Or Mid(STR, I, 1) = "_" Then
        STR = Mid(STR, 1, I) & UCase(Mid(STR, I + 1, 1)) & Mid(STR, I + 2)
        I = I + 1
    Else
        STR = Mid(STR, 1, I - 1) & LCase(Mid(STR, I, 1)) & Mid(STR, I + 1)
    End If
Next
     
    ProperCase = STR
End Function

Public Function CreateSystemDSN(CompanyName As String)
    CompanyName = RemoveSpaces(CompanyName)
    CompanyName = UCase(CompanyName)
    Dim UserName As String
    Dim Password As String
    UserName = "USRPVM" & UCase(CompanyName)
    Password = "PWDPVM" & UCase(CompanyName)
    modDSN.CreateDSN DriverName, CompanyName, ServerName, DatabaseName, UserName, Password
End Function

Private Function Sender()
    Dim M_NME, M_ADR, M_TFX, M_WDS As String
    Dim RST_DIR As New ADODB.Recordset
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR  WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'TT1'", CON, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_NME = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR  WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'TT2'", CON, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_ADR = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR  WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'TT3'", CON, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_TFX = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR  WHERE DIR_IDY = 'WDS' AND DIR_IDX = 'IDY'", CON, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_WDS = Trim(RST_DIR("DIR_ALS"))
    End If

    If Dir(SetupFolder & "\DST_INF.DAT") = "" Then
        Set F_VER = FS_SOU.OpenTextFile(SetupFolder & "\DST_INF.DAT", 2, True)
        F_VER.WriteLine "Distributor : " & M_NME
        F_VER.WriteLine "Address     : " & M_ADR
        F_VER.WriteLine "Telephone   : " & M_TFX
        F_VER.WriteLine "DAS-CS Code-ID : " & M_WDS
        F_VER.Close
    End If

End Function

Public Function MailSend(sSendTo As String, sSubject As String, sText As String, Optional sAttachFile As String) As Boolean
'Dim MAPISession1 As MAPISession
'Dim MAPIMessages1 As MAPIMessages
    On Error GoTo ErrHandler
    With MAPISession1
        If .SessionID = 0 Then
            .DownLoadMail = False
            .LogonUI = True
            .SignOn
            .NewSession = True
            MAPIMessages1.SessionID = .SessionID
        End If
    End With
    With MAPIMessages1
        .Compose
        .RecipAddress = sSendTo
        .AddressResolveUI = True
        .ResolveName
        .MsgSubject = sSubject
        .MsgNoteText = sText
        If Len(sAttachFile) > 0 And Len(Dir$(sAttachFile)) > 0 Then
            .AttachmentPathName = sAttachFile
        Else
            .AttachmentCount = 0
        End If
        .Send False
    End With
    MailSend = True
    Exit Function
ErrHandler:
    Debug.Print err.Description
    MailSend = False
End Function

Public Function ItemPurchaseCalculation(CurrentLevel As Double, PurchaseRate As Double, discount As Double, Tax As Double, SubUnit As Double) As Double

Dim DISAMT As Double
Dim TaxAmt As Double
Dim TotalAmt As Double
Dim CurrentVal As Double
CurrentVal = ((CurrentLevel / SubUnit) + ((CurrentLevel Mod SubUnit) / 1000)) * PurchaseRate
DISAMT = (CurrentVal * discount) / 100
TaxAmt = ((CurrentVal - DISAMT) * Tax) / 100
TotalAmt = CurrentVal - DISAMT + TaxAmt
ItemPurchaseCalculation = TotalAmt
End Function



Public Function InsertIntoCLD(TagValue As String, DT As Date)
Dim rs As New Recordset
Dim SqlStatement As String
Dim RS1 As New Recordset
CON.Execute "DELETE FROM CLD"
SqlStatement = "INSERT INTO CLD (SELECT DISTINCT BAT_IDY  ,0,0,0,0,0,0,0,0 FROM inv)"
CON.Execute SqlStatement

If TagValue = "COK_STK_BAT" Or TagValue = "COK_STK_BRD" Or TagValue = "COK_STK_SEG" Then
    RS1.Open "SELECT SUM(PRD_UOM),BAT_IDY AS PRD_IDY,INV.TXN_TYP,INV.STK_TYP FROM INV, INL WHERE INL.INV_IDY = INV.INV_IDY AND INL.NET_AMT > 0 AND INV.STK_UPD = 'Y' AND INV.INV_DAT < '" & Format(DT, modGeneral.OraDtFormat) & "' GROUP BY BAT_IDY, INV.TXN_TYP, INV.STK_TYP", CON, adOpenKeyset, adLockPessimistic 'Modified by Radhika and Ramanesh
    TagValue = "PRD_UOM"
Else
    TagValue = "TOT_PCS"
    RS1.Open "SELECT SUM(tot_pcs),BAT_IDY AS PRD_IDY,INV.TXN_TYP,INV.STK_TYP FROM INV, INL WHERE INL.INV_IDY = INV.INV_IDY AND INL.NET_AMT > 0 AND INV.STK_UPD = 'Y' AND INV.INV_DAT < '" & Format(DT, modGeneral.OraDtFormat) & "' GROUP BY BAT_IDY, INV.TXN_TYP, INV.STK_TYP", CON, adOpenKeyset, adLockPessimistic 'Modified by Radhika and Ramanesh
'    RS1.Open "select sum(tot_pcs),BAT_IDY AS PRD_IDY,txn_typ,stk_typ from inv where stk_upd = 'Y' AND inv_dat < '" & Format(DT, modGeneral.OraDtFormat) & "' group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic
End If

'SqlStatement = "UPDATE CLD C SET C.SAL_STK = C.SAL_STK + (SELECT SUM(" & TagValue & ") FROM INV WHERE TXN_TYP ='SAL' AND (STK_TYP = 'Salable' OR STK_TYP='Free') AND C.PRD_IDY = INV.BAT_IDY AND STK_UPD = 'Y' AND INV_DAT <'" & Format(dt, modGeneral.OraDtFormat) & "' GROUP BY BAT_IDY, C.PRD_IDY)"
'con.Execute SqlStatement
'SqlStatement = "UPDATE CLD C SET C.SAL_STK = C.SAL_STK - (SELECT SUM(" & TagValue & ") FROM INV WHERE TXN_TYP ='SAL' AND (STK_TYP = 'TakeBack' OR STK_TYP='Damage' OR STK_TYP ='Leakage') AND C.PRD_IDY = INV.BAT_IDY AND STK_UPD = 'Y' AND INV_DAT <'" & Format(dt, modGeneral.OraDtFormat) & "' GROUP BY BAT_IDY)"
'con.Execute SqlStatement
'SqlStatement = "UPDATE CLD C SET C.OPN_STK = C.OPN_STK + (SELECT SUM(" & TagValue & ") FROM INV WHERE TXN_TYP ='OPS' AND C.PRD_IDY = INV.BAT_IDY AND STK_UPD = 'Y' AND INV_DAT <'" & Format(dt, modGeneral.OraDtFormat) & "' GROUP BY BAT_IDY)"
'con.Execute SqlStatement
'SqlStatement = "UPDATE CLD C SET C.PUR_STK = C.PUR_STK + (SELECT SUM(" & TagValue & ") FROM INV WHERE TXN_TYP ='PUR' AND C.PRD_IDY = INV.BAT_IDY AND STK_UPD = 'Y' AND INV_DAT <'" & Format(dt, modGeneral.OraDtFormat) & "' GROUP BY BAT_IDY)"
'con.Execute SqlStatement
'SqlStatement = "UPDATE CLD C SET C.PRT_STK = C.PRT_STK + (SELECT SUM(" & TagValue & ") FROM INV WHERE TXN_TYP ='PRT' AND C.PRD_IDY = INV.BAT_IDY AND STK_UPD = 'Y' AND INV_DAT <'" & Format(dt, modGeneral.OraDtFormat) & "' GROUP BY BAT_IDY)"
'con.Execute SqlStatement
'SqlStatement = "UPDATE CLD C SET C.STI_STK = C.STI_STK + (SELECT SUM(" & TagValue & ") FROM INV WHERE TXN_TYP ='STI' AND C.PRD_IDY = INV.BAT_IDY AND STK_UPD = 'Y' AND INV_DAT <'" & Format(dt, modGeneral.OraDtFormat) & "' GROUP BY BAT_IDY)"
'con.Execute SqlStatement
'SqlStatement = "UPDATE CLD C SET C.SRN_STK = C.SRN_STK + (SELECT SUM(" & TagValue & ") FROM INV WHERE TXN_TYP ='SRN' AND C.PRD_IDY = INV.BAT_IDY AND STK_UPD = 'Y' AND INV_DAT <'" & Format(dt, modGeneral.OraDtFormat) & "' GROUP BY BAT_IDY)"
'con.Execute SqlStatement
'SqlStatement = "UPDATE CLD C SET C.STO_STK = C.STO_STK + (SELECT SUM(" & TagValue & ") FROM INV WHERE TXN_TYP ='STO' AND C.PRD_IDY = INV.BAT_IDY AND STK_UPD = 'Y' AND INV_DAT <'" & Format(dt, modGeneral.OraDtFormat) & "' GROUP BY BAT_IDY)"
'con.Execute SqlStatement
While Not RS1.EOF
    If rs.State = 1 Then rs.Close
    SqlStatement = "SELECT COUNT(*) FROM CLD " & " WHERE PRD_IDY = '" & Trim(RS1!prd_idy) & "'"
    rs.Open SqlStatement, CON, adOpenKeyset, adLockPessimistic
    If rs(0) <= 0 Then
    End If
    If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
        CON.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY = '" & Trim(RS1!prd_idy) & "'", d
    ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
        CON.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!prd_idy) & "'", d
    ElseIf RS1!TXN_TYP = "OPS" Then
        CON.Execute "UPDATE CLD SET OPN_STK = " & Val(RS1(0)) & " WHERE TRIM(PRD_IDY) = '" & Trim(RS1!prd_idy) & "'", d
    ElseIf RS1!TXN_TYP = "PUR" Then
        CON.Execute "UPDATE CLD SET PUR_STK = " & Val(RS1(0)) & " WHERE TRIM(PRD_IDY) = '" & Trim(RS1!prd_idy) & "'", d
    ElseIf RS1!TXN_TYP = "PRT" Then
        CON.Execute "UPDATE CLD SET PRT_STK = " & Val(RS1(0)) & " WHERE TRIM(PRD_IDY) = '" & Trim(RS1!prd_idy) & "'", d
    ElseIf RS1!TXN_TYP = "STI" Then
        CON.Execute "UPDATE CLD SET STI_STK = " & Val(RS1(0)) & " WHERE TRIM(PRD_IDY) = '" & Trim(RS1!prd_idy) & "'", d
    ElseIf RS1!TXN_TYP = "SRN" Then
        CON.Execute "UPDATE CLD SET SRN_STK = " & Val(RS1(0)) & " WHERE TRIM(PRD_IDY) = '" & Trim(RS1!prd_idy) & "'", d
    ElseIf RS1!TXN_TYP = "STO" Then
        CON.Execute "UPDATE CLD SET STO_STK = " & Val(RS1(0)) & " WHERE TRIM(PRD_IDY) = '" & Trim(RS1!prd_idy) & "'", d
    End If
    RS1.MoveNext
Wend
    CON.Execute "UPDATE CLD SET CLS_STK= OPN_STK + PUR_STK - SAL_STK + SRN_STK - PRT_STK + STI_STK - STO_STK"

End Function

Public Sub FastBackup()
On Error Resume Next
    If MsgBox("Make Fast Backup of DataFiles ?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
        ChDir EFF_LIB1.CompanyFolderPath
        FST_BKP
    End If
End Sub

Public Sub FST_BKP()
M_MDI.MousePointer = vbHourglass
Dim DIR_RS As New ADODB.Recordset
  If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", CON, adOpenKeyset, adLockOptimistic
    If DIR_RS.RecordCount > 0 Then DIR_RS.MoveFirst
       DIR_RS("DIR_MSG") = Trim(M_WDS_SRL)
       DIR_RS.Update
       M_WDS_IDY = Trim(DIR_RS!DIR_ALS)
    If DIR_RS.State = 1 Then DIR_RS.Close
    'Zip File Creation
            Dim DEX As String
            DEX = "C:\DEX"
            If Dir(DEX, vbDirectory) = "" Then
                MkDir DEX
            End If
            Dim FS_SOU, F_VER As Object
            Set FS_SOU = CreateObject("Scripting.filesystemobject")
            
            If Dir(App.Path & "\VER.BAT") = "" Then
                Set F_VER = FS_SOU.OpenTextFile(App.Path & "\VER.BAT", 2, True)
                F_VER.WriteLine "DIR " & App.Path & "\*.* /OG > VER_APP.TXT"
                F_VER.WriteLine "DIR " & CurDir & "\*.* /OG > VER_DAT.TXT"
                F_VER.Close
            Else
                CopyFile App.Path & "\VER.BAT", CurDir & "\VER.BAT", 0
            End If
            
            If Dir(CurDir & "\VER.BAT") <> "" Then
                Shell CurDir & "\VER.BAT", vbHide
            End If
            
            Dim M_NME, M_ADR, M_TFX, M_WDS As String
            Dim RST_DIR As New ADODB.Recordset
            
            If RST_DIR.State = 1 Then RST_DIR.Close
            RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TT1'", CON, adOpenKeyset, adLockOptimistic
            If RST_DIR.EOF = False Then
                M_NME = Trim(RST_DIR("DIR_MSG"))
            End If
            
            If RST_DIR.State = 1 Then RST_DIR.Close
            RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TT2'", CON, adOpenKeyset, adLockOptimistic
            If RST_DIR.EOF = False Then
                M_ADR = Trim(RST_DIR("DIR_MSG"))
            End If
            
            If RST_DIR.State = 1 Then RST_DIR.Close
            RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TT3'", CON, adOpenKeyset, adLockOptimistic
            If RST_DIR.EOF = False Then
                M_TFX = Trim(RST_DIR("DIR_MSG"))
            End If
            
            If RST_DIR.State = 1 Then RST_DIR.Close
            RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'WDS' AND DIR_IDX LIKE 'IDY'", CON, adOpenKeyset, adLockOptimistic
            If RST_DIR.EOF = False Then
                M_WDS = Trim(RST_DIR("DIR_ALS"))
            End If

            If Dir(CurDir & "\DIR_INF.TXT") = "" Then
                Set F_VER = FS_SOU.OpenTextFile(CurDir & "\DIR_INF.TXT", 2, True)
                F_VER.WriteLine "Distributor : " & M_NME
                F_VER.WriteLine "Address     : " & M_ADR
                F_VER.WriteLine "Telephone   : " & M_TFX
                F_VER.WriteLine "DAS Code-ID : " & M_WDS
                F_VER.Close
            End If
            If Dir(CurDir & "\CHGDBF.BAT") = "" Then
                MsgBox "Rename (CHGDBF.BAT) DBF Utility Batch File Not Found ! Quitting Application !", vbCritical, "DAS Version 4.0"
                End
            Else
                Shell "CHGDBF.BAT", vbMinimizedFocus
            End If
            M_FIL_NME = "\" & "FBKPDAS" & Trim(M_WDS_IDY) & Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & ".ZIP"
            Dim Myzip As ColaZip.GSColaZip32
            Set Myzip = New ColaZip.GSColaZip32
            Myzip.RootDirectory = "\"
            Myzip.ZipFileName = "C:\DEX" & M_FIL_NME
            Myzip.AddFile (CurDir & "\EFF.*")
            Myzip.AddFile (CurDir & "\*.DBF")
            Myzip.AddFile (CurDir & "\*.FPT")
            Myzip.AddFile (CurDir & "\VER_APP.TXT")
            Myzip.AddFile (CurDir & "\VER_DAT.TXT")
            Myzip.AddFile (CurDir & "\DIR_INF.TXT")
            Myzip.AddFile (CurDir & "\SHUTDOWN.LOG")
'            Myzip.AddFile (curdir & "\HHT\DAS_STOCK.DAT")
            'Myzip.AddFile ("C:\DEX\OutBox\*.DAT")
            
            If M_MDI.SysInfo1.OSPlatform = 1 Then
                Myzip.AddFile "c:\windows\system\EFFMCG4.txt"
            ElseIf M_MDI.SysInfo1.OSPlatform = 2 Then
                Myzip.AddFile "c:\winnt\system\EFFMCG4.txt"
            Else
                Myzip.AddFile "c:\windows\system\EFFMCG4.txt"
            End If
            
            Myzip.MakeColaZip ("bkp")   'creation of Zip File with Password bkp (in lower case)
            
            If Dir(CurDir & "\MAKDBF.BAT") = "" Then
                MsgBox "Rename (MAKDBF.BAT) DBF Utility Batch File Not Found ! Quitting Application !", vbCritical, "DAS Version 4.0"
                
                End
            Else
                Shell "MAKDBF.BAT", vbMinimizedFocus
            End If
            
            MsgBox ("Zip File " & M_FIL_NME & " Created in the Folder " & DEX), vbInformation, "DAS Version 4.0"
            ' Zip File Creation - Completed
            M_MDI.MousePointer = vbDefault
            
            
            M_MSG_STR = "Zipped Data From : " & Trim(M_NME) & " , " & Trim(M_ADR) & " , " & Trim(M_TFX) & " , " & Trim(M_WDS)
            M_MSG_EML = EFF_LIB1.GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'EML' AND DIR_IDX LIKE 'IDY'")
            M_MSG_SUB = "Zip Data Xfer Fast Backup - Data Extract : " & M_WDS
            
            ' Zip File Send by eMail Option
            M_MAIL_ATTMNT = "C:\DEX" & Trim(M_FIL_NME)
            If MsgBox("Send Zip File Data to Help Desk Now ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
                If Trim(M_MSG_EML) <> "" Then
                    MailSend Trim(M_MSG_EML), Trim(M_MSG_SUB), Trim(M_MSG_STR), Trim(M_MAIL_ATTMNT)
                Else
                    MailSend "effmcg@consoftsystems.com", Trim(M_MSG_SUB), Trim(M_MSG_STR), Trim(M_MAIL_ATTMNT)
                End If
                MsgBox ("Data Send via MAPI Executed!"), vbInformation, "DAS Version 4.0"
                Command2.SetFocus
'                Unload Me
            Else
                Command2.SetFocus
            End If
End Sub

Public Sub RepairDatabase()
    Text = EFF_LIB1.CompanyFolderPath
    CurDir Text
    If MsgBox("Repair Selected Company Files?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
        ChDir Text
        If Dir(Text & "\REPAIR.BAT") = "" Then
            MsgBox "Cannot Send Commands, File REPAIR.BAT Not Found!", vbExclamation, "DAS Version 4.0"
        Else
            Shell Text & "\REPAIR.BAT", vbHide
        End If
    End If
    ReSetConnectionObjects EFF_LIB1.CompanyName
End Sub

Public Sub IndexData()
    CloseConnectionObjects
    Set OraConnection = Nothing
    Set CON = Nothing
On Error Resume Next
    Text = EFF_LIB1.CompanyFolderPath
    CurDir Text
    If MsgBox("Index Database Tables ?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
        ChDir Text
        If Dir(Text & "\ind.exe") = "" Then
            MsgBox "Cannot Index, Index File IND.EXE Not Found!", vbExclamation, "DAS Version 4.0"
        Else
            Shell Text & "\ind.exe", vbMaximizedFocus
        End If
    End If
    'ReSetConnectionObjects EFF_LIB1.CompanyName

End Sub

