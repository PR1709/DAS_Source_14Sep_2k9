Dim CON As New ADODB.Connection
CON.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CurDir & "\EFF.DBC;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
Me.MousePointer = vbHourglass
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
            RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' .AND. DIR_IDX LIKE 'TT1'", CON, adOpenKeyset, adLockOptimistic
            If RST_DIR.EOF = False Then
                M_NME = Trim(RST_DIR("DIR_MSG"))
            End If
            
            If RST_DIR.State = 1 Then RST_DIR.Close
            RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' .AND. DIR_IDX LIKE 'TT2'", CON, adOpenKeyset, adLockOptimistic
            If RST_DIR.EOF = False Then
                M_ADR = Trim(RST_DIR("DIR_MSG"))
            End If
            
            If RST_DIR.State = 1 Then RST_DIR.Close
            RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' .AND. DIR_IDX LIKE 'TT3'", CON, adOpenKeyset, adLockOptimistic
            If RST_DIR.EOF = False Then
                M_TFX = Trim(RST_DIR("DIR_MSG"))
            End If
            
            If RST_DIR.State = 1 Then RST_DIR.Close
            RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'WDS' .AND. DIR_IDX LIKE 'IDY'", CON, adOpenKeyset, adLockOptimistic
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
            
            If CMP_SEL.SysInfo1.OSPlatform = 1 Then
                Myzip.AddFile "c:\windows\system\EFFMCG4.txt"
            ElseIf CMP_SEL.SysInfo1.OSPlatform = 2 Then
                Myzip.AddFile "c:\winnt\system\EFFMCG4.txt"
            Else
                Myzip.AddFile "c:\windows\system\EFFMCG4.txt"
            End If
            
            Myzip.MakeColaZip ("bkp")   'creation of Zip File with Password bkp (in lower case)
            MsgBox ("Zip File " & M_FIL_NME & " Created in the Folder " & DEX), vbInformation, "DAS Version 4.0"
            ' Zip File Creation - Completed
            Me.MousePointer = vbDefault
            
            
            M_MSG_STR = "Zipped Data From : " & Trim(M_NME) & " , " & Trim(M_ADR) & " , " & Trim(M_TFX) & " , " & Trim(M_WDS)
            M_MSG_EML = GetValue("DIR_MSG", "DIR WHERE DIR_IDY LIKE 'EML' AND DIR_IDX LIKE 'IDY'")
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
'                Command2.SetFocus
'                Unload Me
            Else
'                Command2.SetFocus
            End If
End Sub