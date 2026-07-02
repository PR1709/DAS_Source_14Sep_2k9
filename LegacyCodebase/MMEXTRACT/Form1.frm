VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Dim con As ADODB.Connection
Dim PNL As ADODB.Connection

Dim sqlStatement As String
Dim REC As Recordset
Dim FromDate As Date
Dim ToDate As Date
Dim rs As New Recordset
Dim RS1 As New Recordset

Private Sub MM_EX_FILE()
    On Error Resume Next
    Dim fpath As String
    fpath = "C:\DEX\INBOX"
    Dim fso As New FileSystemObject
    If fso.FileExists(fpath & "\MM.DBF") = False Then
        MsgBox "File Not Found - " & fpath & "\MM.DBF"
        Exit Sub
    End If
    Set REC = New Recordset
    sqlStatement = "SELECT * FROM MM"
    
    Set con = New Connection
    Set PNL = New Connection
    con.Open "Driver={Microsoft Visual Foxpro Driver};UID=;PWD=;SourceDB=" & App.Path & "\Eff.dbc;SourceType=DBC;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    PNL.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & fpath & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
    
    REC.Open sqlStatement, PNL, adOpenKeyset, adLockPessimistic
    FromDate = REC!frm_dat
    ToDate = REC!to_dat
    
    
    
' * Distributor ID Identification
        Dim DIR_RS As New ADODB.Recordset
        Dim M_WDS_IDY As String
        Dim M_WDS_SRL As String
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            If DIR_RS.BOF = False Then
                DIR_RS.MoveFirst
                M_WDS_IDY = Trim(DIR_RS("DIR_ALS"))
            End If
            If Len(M_WDS_IDY) < 8 Then
                MsgBox ("Distributor ID Required for this Process! Obtain Distributor ID and Start Again"), vbCritical, "DAS Version 4.0"
                Exit Sub
            End If
            M_WDS_SRL = (Val(DIR_RS("DIR_MSG")) + 1)
            If Len(Trim(M_WDS_SRL)) < 3 Then
                 While Len(Trim(M_WDS_SRL)) < 3
                      M_WDS_SRL = "0" & M_WDS_SRL
                 Wend
            End If
        End If
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'SYS' and DIR_IDX = 'VER'", con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            If DIR_RS.BOF = False Then
                DIR_RS.MoveFirst
                DIR_RS("DIR_MSG") = Trim("01.01.01")
                DIR_RS.Update
            End If
        End If
' * Distributor ID Identification
        Dim Pending_RS As New ADODB.Recordset
        If Pending_RS.State = 1 Then Pending_RS.Close
        Pending_RS.Open "select * from DIR where DIR_DAT >= CTOD('" & Format(FromDate, "DD/MM/YYYY") & "') AND DIR_DAT <= CTOD('" & Format(ToDate, "DD/MM/YYYY") & "') AND alltrim(DIR_MSG) = 'Y' AND substr(alltrim(DIR_IDY),1,3) = 'DAY'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh
        'SELECT * FROM DIR WHERE CDOW(DIR_DAT) = "Friday" AND TRIM(dir_msg) = "Y"
        If Pending_RS.RecordCount > 0 Then
            Pending_RS.MoveFirst
            ST_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
            Pending_RS.MoveLast
            EN_DAT = CDate(Trim(Pending_RS("DIR_DAT")))
        End If

'If MsgBox("Create Data File for Data Transfer Upto " & Format(ToDate, "DD/MM/YYYY"), vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
        
    ' * : Beginning of DAS_CUSTOMER.DAT Creation Code for Data Transfer
   ' Me.MousePointer = vbHourglass
    Dim F_OMS, F_LOC, F_MAR, F_SMN
    Dim FS_SOU, FS_LOC, FS_MAR, FS_SMN As Object
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    Set FS_LOC = CreateObject("Scripting.filesystemobject")
    Set FS_MAR = CreateObject("SCripting.filesystemobject")
    Set FS_SMN = CreateObject("SCripting.filesystemobject")
    
    If FS_SOU.FolderExists(CurDir & "\HHT") = False Then
       FS_SOU.CreateFolder (CurDir & "\HHT")
    End If
    
    If FS_SOU.FileExists(CurDir & "\HHT\DAS_CUSTOMER.DAT") = True Then
        FS_SOU.DeleteFile (CurDir & "\HHT\DAS_CUSTOMER.DAT")
    End If
'    If FS_LOC.FileExists(CURDIR & "\HHT\LOCATION.TXT") = True Then
'        FS_LOC.DeleteFile (CURDIR & "\HHT\LOCATION.TXT")
'    End If
'    If FS_MAR.FileExists(CURDIR & "\HHT\MARKET.TXT") = True Then
'        FS_MAR.DeleteFile (CURDIR & "\HHT\MARKET.TXT")
'    End If
'    If FS_SMN.FileExists(CURDIR & "\HHT\SALESMAN.TXT") = True Then
'        FS_SMN.DeleteFile (CURDIR & "\HHT\SALESMAN.TXT")
'    End If
    
    Set F_OMS = FS_SOU.OpenTextFile(CurDir & "\HHT\DAS_CUSTOMER.DAT", 2, True)
'    Set F_LOC = FS_LOC.OpenTextFile(CURDIR & "\HHT\LOCATION.TXT", 2, True)
'    Set F_MAR = FS_MAR.OpenTextFile(CURDIR & "\HHT\MARKET.TXT", 2, True)
'    Set F_SMN = FS_SMN.OpenTextFile(CURDIR & "\HHT\SALESMAN.TXT", 2, True)
    
    Dim J As Integer
    J = 0
    
    
    Dim Category_RS As New ADODB.Recordset
    Dim Location_RS As New ADODB.Recordset
    Dim Market_RS As New ADODB.Recordset
    Dim Salesman_RS As New ADODB.Recordset
    Dim Invoice_Chk_RS As New ADODB.Recordset
    
    Dim Retailer_RS As New ADODB.Recordset
    If Retailer_RS.State = 1 Then Retailer_RS.Close
    Retailer_RS.Open "select * from ACT where ACT_TYP = 'SAL' or ACT_TYP = 'PUR'", con, adOpenKeyset, adLockPessimistic  'Modified by Ramanesh  CST_TYP <> 'Van'
    
    If Retailer_RS.RecordCount > 0 Then Retailer_RS.MoveFirst
       
       Dim L_CAT_NME As String
       Dim L_LOC_NME As String
       Dim L_MAR_NME As String
       Dim L_DST_NME As String
       Dim L_MAR_GRD As String
       
       If Category_RS.State = 1 Then Category_RS.Close
       Category_RS.Open "SELECT * FROM CAT", con, adOpenKeyset, adLockPessimistic
       If Location_RS.State = 1 Then Location_RS.Close
       Location_RS.Open "SELECT * FROM LOC", con, adOpenKeyset, adLockPessimistic
       If Market_RS.State = 1 Then Market_RS.Close
       Market_RS.Open "SELECT * FROM MAR", con, adOpenKeyset, adLockPessimistic
       
       While Not Retailer_RS.EOF
       
          If Category_RS.RecordCount > 0 Then Category_RS.MoveFirst
          Category_RS.Find "CAT_IDY='" & Trim(Retailer_RS("CAT_IDY")) & "'"
          If Category_RS.EOF = False Then
             L_CAT_NME = Category_RS("CAT_NME")
          End If
          
          If Location_RS.RecordCount > 0 Then Location_RS.MoveFirst
          Location_RS.Find "LOC_IDY='" & Trim(Retailer_RS("LOC_IDY")) & "'"
          If Location_RS.EOF = False Then
             L_LOC_NME = Location_RS("LOC_NME")
          End If
          
          If Market_RS.RecordCount > 0 Then Market_RS.MoveFirst
          Market_RS.Find "MAR_IDY='" & Trim(Retailer_RS("MAR_IDY")) & "'"
          If Market_RS.EOF = False Then
             L_MAR_NME = Market_RS("MAR_NME")
             L_DST_NME = Market_RS("MAR_PGR")
             L_MAR_GRD = Market_RS("MAR_GRD")
             L_MAR_CSG = Market_RS("MAR_CSG")
             L_MAR_CLU = Market_RS("MAR_CLU")
             L_MAR_ROC = Market_RS("MAR_ROC")
             L_MAR_TWC = Market_RS("MAR_TWC")
             L_MAR_TWN = Market_RS("MAR_TWN")
             L_MAR_DST = Market_RS("MAR_DST")
             L_MAR_STA = Market_RS("MAR_STA")
             L_MAR_COU = Market_RS("MAR_COU")
          End If
          
'          F_OMS.WriteLine Trim(M_WDS_IDY) & "|" & Trim(Retailer_RS("ACT_IDY")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(L_DST_NME) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(L_LOC_NME) & "|" & Trim(L_CAT_NME) & "|" & Trim(L_MAR_NME) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("DBT_TYP"))
          If Trim(Retailer_RS("ACT_TYP")) = "PUR" Then
                'F_OMS.WriteLine Trim(Retailer_RS("ACT_IDY")) & "|" & "S" & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(Retailer_RS("LOC_IDY")) & "|" & Trim(Retailer_RS("MAR_IDY")) & "|" & Trim(Retailer_RS("CAT_IDY")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("DBT_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("ROU_IDY"))
                'J = J + 1
          Else
                'F_OMS.WriteLine Trim(Retailer_RS("ACT_IDY")) & "|" & "C" & "|" & Trim(Retailer_RS("CST_TYP")) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_AD3")) & "|" & Trim(Retailer_RS("ACT_AD4")) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_MBL")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(Retailer_RS("ACT_CPN")) & "|" & Trim(Retailer_RS("LOC_IDY")) & "|" & Trim(Retailer_RS("MAR_IDY")) & "|" & Trim(Retailer_RS("CAT_IDY")) & "|" & Trim(Retailer_RS("RET_TYP")) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(Retailer_RS("DBT_TYP")) & "|" & Trim(Retailer_RS("VAN_CAP")) & "|" & Trim(Retailer_RS("ROU_IDY"))
                If Invoice_Chk_RS.State = 1 Then Invoice_Chk_RS.Close
                Invoice_Chk_RS.Open "SELECT * FROM INL WHERE CST_IDY = '" & Trim(Retailer_RS("ACT_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                If Invoice_Chk_RS.RecordCount > 0 Then
                    F_OMS.WriteLine Trim(M_WDS_IDY) & Right(Trim(Retailer_RS("ACT_IDY")), 4) & "|" & Trim(Retailer_RS("ACT_NME")) & "|" & "I" & "|" & Trim(Retailer_RS("ACT_AD1")) & "|" & Trim(Retailer_RS("ACT_AD2")) & "|" & Trim(Retailer_RS("ACT_ZIP")) & "|" & Trim(Retailer_RS("ACT_TFX")) & "|" & Trim(Retailer_RS("ACT_FAX")) & "|" & Trim(Retailer_RS("ACT_EML")) & "|" & Trim(L_MAR_COU) & "|" & Trim(L_MAR_STA) & "|" & Trim(L_MAR_DST) & "|" & Trim(L_MAR_TWN) & "|" & Trim(L_CAT_NME) & "|" & Trim(L_MAR_CSG) & "|" & Trim(Retailer_RS("BBN_ONE")) & "|" & Trim(L_MAR_CLU) & "|" & Trim(L_MAR_ROC) & "|" & Trim(L_MAR_TWC) & "|" & Trim(Retailer_RS("DBT_TYP"))
                    J = J + 1
                End If
          End If
          Retailer_RS.MoveNext
       Wend
'''       MsgBox "" & J & ": Outlet Records Created for Data Transfer!"
'''    Else
'''       MsgBox ("No Outlet Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
'''       cancel
'''    End If
'       Dim n As Integer
'       n = 0
'       If Location_RS.State = 1 Then Location_RS.Close
'       Location_RS.Open "SELECT * FROM LOC", CON, adOpenKeyset, adlockpessimistic
'       While Not Location_RS.EOF
'          F_LOC.WriteLine Trim(Location_RS("LOC_IDY")) & "|" & Trim(Location_RS("LOC_NME"))
'          Location_RS.MoveNext
'          n = n + 1
'       Wend
'
'       Dim o As Integer
'       o = 0
'       If Market_RS.State = 1 Then Market_RS.Close
'       Market_RS.Open "SELECT * FROM MAR", CON, adOpenKeyset, adlockpessimistic
'       While Not Market_RS.EOF
'          F_MAR.WriteLine Trim(Market_RS("MAR_IDY")) & "|" & Trim(Market_RS("MAR_NME")) & "|" & Trim(Market_RS("MAR_PGR")) & "|" & Trim(Market_RS("MAR_GRD"))
'          Market_RS.MoveNext
'          o = o + 1
'       Wend
'
'       Dim P As Integer
'       P = 0
'       If Salesman_RS.State = 1 Then Salesman_RS.Close
'       Salesman_RS.Open "SELECT * FROM SMN", CON, adOpenKeyset, adlockpessimistic
'       While Not Salesman_RS.EOF()
'          F_SMN.WriteLine Trim(Salesman_RS("SMN_IDY")) & "|" & Trim(Salesman_RS("SMN_NME"))
'          Salesman_RS.MoveNext
'          P = P + 1
'       Wend
'
        F_OMS.Close
'        F_LOC.Close
'        F_MAR.Close
'        F_SMN.Close
'    cancel
    
    ' * : End of DAS_CUSTOMER.DAT Creation Code for Data Transfer
    
    ' * : Beginning of DAS_SECONDARY.DAT Creation Code for Data Transfer
    
    Dim F_INV
    Dim F_INL
    Dim F_VAN
    Dim F_EIS
    
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    Set FS_INL = CreateObject("Scripting.filesystemobject")
'    Set FS_VAN = CreateObject("Scripting.filesystemobject")
'    Set FS_EIS = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CurDir & "\HHT") = False Then
       FS_SOU.CreateFolder (CurDir & "\HHT")
    End If
    
    If FS_SOU.FileExists(CurDir & "\HHT\DAS_SECONDARY.DAT") = True Then
        FS_SOU.DeleteFile (CurDir & "\HHT\DAS_SECONDARY.DAT")
    End If
    
    If FS_INL.FileExists(CurDir & "\HHT\DAS_PURCHASE.DAT") = True Then
        FS_INL.DeleteFile (CurDir & "\HHT\DAS_PURCHASE.DAT")
    End If
    
'    If FS_VAN.FileExists(CURDIR & "\HHT\LOAD_SHEET.TXT") = True Then
'        FS_VAN.DeleteFile (CURDIR & "\HHT\LOAD_SHEET.TXT")
'    End If
'
'    If FS_EIS.FileExists(CURDIR & "\HHT\EMPTY_ISSUE.TXT") = True Then
'        FS_EIS.DeleteFile (CURDIR & "\HHT\EMPTY_ISSUE.TXT")
'    End If
    
    Set F_INV = FS_SOU.OpenTextFile(CurDir & "\HHT\DAS_SECONDARY.DAT", 2, True)
    Set F_INL = FS_INL.OpenTextFile(CurDir & "\HHT\DAS_PURCHASE.DAT", 2, True)
'    Set F_VAN = FS_VAN.OpenTextFile(CURDIR & "\HHT\LOAD_SHEET.TXT", 2, True)
'    Set F_EIS = FS_EIS.OpenTextFile(CURDIR & "\HHT\EMPTY_ISSUE.TXT", 2, True)
    
    Dim K As Double
    Dim r As Double
    Dim s As Double
    Dim T As Double
    K = 0
    r = 0
    s = 0
    T = 0
    Dim INL_RS As New ADODB.Recordset
    Dim INV_RS As New ADODB.Recordset
    Dim LNV_RS As New ADODB.Recordset
    Dim PRD_RS As New ADODB.Recordset
    Dim BAT_RS As New ADODB.Recordset
    Dim SCH_RS As New ADODB.Recordset
    Dim VAN_RS As New ADODB.Recordset
    Dim SNV_RS As New ADODB.Recordset
    Dim ACT_RS As New ADODB.Recordset
    
    Dim L_CST_IDY As String
    Dim L_ORD_IDY As String
    Dim L_BAT_NME As String
    Dim L_PRD_COD As String
    Dim L_MER_NUM As String
    
    If INL_RS.State = 1 Then INL_RS.Close
    If INV_RS.State = 1 Then INV_RS.Close
    If LNV_RS.State = 1 Then LNV_RS.Close
    If PRD_RS.State = 1 Then PRD_RS.Close
    If BAT_RS.State = 1 Then BAT_RS.Close
    If SCH_RS.State = 1 Then SCH_RS.Close
       
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * FROM DIR where substr(alltrim(DIR_IDY),1,3) = 'DAY' AND alltrim(DIR_MSG) = 'Y' AND DIR_DAT >= CTOD('" & FromDate & "') AND DIR_DAT <= CTOD('" & ToDate & "')", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveFirst
        ST_DAT = DIR_RS("DIR_DAT")
        BEG_DAT = DIR_RS("DIR_DAT")
        DIR_RS.MoveLast
        EN_DAT = DIR_RS("DIR_DAT")
        END_DAT = DIR_RS("DIR_DAT")
    End If

    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    INV_RS.Open "SELECT * FROM INV WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
    BAT_RS.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
    SCH_RS.Open "SELECT * FROM SCH", con, adOpenKeyset, adLockPessimistic
    
    If INV_RS.RecordCount > 0 Then INV_RS.MoveFirst
    
        While Not INV_RS.EOF

            If PRD_RS.RecordCount > 0 Then PRD_RS.MoveFirst
            PRD_RS.Find "PRD_IDY='" & Trim(INV_RS("PRD_IDY")) & "'"
            If PRD_RS.EOF = False Then
               L_PRD_COD = PRD_RS("ALT_IDY")
            End If
            
            If BAT_RS.RecordCount > 0 Then BAT_RS.MoveFirst
            BAT_RS.Find "BAT_IDY='" & Trim(INV_RS("BAT_IDY")) & "'"
            If BAT_RS.EOF = False Then
               L_BAT_NME = BAT_RS("BAT_NME")
               If IsNull(BAT_RS("U_PUR_LPP")) = True Then
                    L_PUR_LPP = 0#
               Else
                    L_PUR_LPP = Val(BAT_RS("U_PUR_LPP"))
               End If
               If IsNull(BAT_RS("U_SAL_LPP")) = True Then
                    L_SAL_LPP = 0#
                    L_RMM_UNT = 0#
               Else
                    L_SAL_LPP = Val(BAT_RS("U_SAL_LPP"))
                    L_RMM_UNT = (Val(BAT_RS("PRD_MRP")) / Val(BAT_RS("SKU_SBU"))) - Val(BAT_RS("U_SAL_LPP"))
               End If
               L_MRP_RAT = (Val(BAT_RS("PRD_MRP")) / Val(BAT_RS("SKU_SBU")))
               If IsNull(BAT_RS("SAL_LPP")) = True Or IsNull(BAT_RS("PUR_LPP")) = True Then
                    L_DMM_UNT = 0#
               Else
                    L_DMM_UNT = Val(BAT_RS("SAL_LPP")) - Val(BAT_RS("PUR_LPP")) 'VALUE TAKEN ON PER CASE BASIS - PNR
               End If
               If IsNull(BAT_RS("P_TOT_VAT")) = True Then
                    PL_TOT_VAT = 0#
                    SL_TOT_VAT = 0#
               Else
                    PL_TOT_VAT = (Val(BAT_RS("P_TOT_VAT")) / Val(BAT_RS("SKU_SBU")))
                    SL_TOT_VAT = (Val(BAT_RS("S_TOT_VAT")) / Val(BAT_RS("SKU_SBU")))
               End If
            End If
            
            If INL_RS.RecordCount > 0 Then INL_RS.MoveFirst
            INL_RS.Find "INV_IDY='" & Trim(INV_RS("INV_IDY")) & "'"
            If INL_RS.EOF = False Then
               L_CST_IDY = INL_RS("CST_IDY")
               L_ORD_IDY = INL_RS("ORD_IDY")
               L_LRN_IDY = INL_RS("LRN_IDY")
               L_LRN_DAT = INL_RS("LRN_DAT")
            End If
            
            If SCH_RS.RecordCount > 0 Then SCH_RS.MoveFirst
            SCH_RS.Find "SCH_IDY='" & Trim(INV_RS("SCH_IDY")) & "'"
            If SCH_RS.EOF = False Then
                L_MER_NUM = SCH_RS("MER_NUM")
            End If
            If BAT_RS.EOF = False Then
                'M_WDS_IDY & "|"Trim(STR(Inv_RS("COL_AMT"))) & "|" &
                If Trim(INV_RS("TXN_TYP")) = "SAL" Then
                    If Trim(INV_RS("STK_TYP")) = "Salable" Then
                        'F_INV.WriteLine Trim(INV_RS("INV_IDY")) & "|" & Format(INV_RS("INV_DAT"), "DD/MM/YYYY") & "|" & Trim(INV_RS("TXN_TYP")) & "|" & Trim(L_CST_IDY) & "|" & Trim(L_PRD_COD) & "|" & QTY(INV_RS("PRD_QTY")) & "|" & AMT(INV_RS("PRD_RAT")) & "|" & Trim(L_BAT_NME) & "|" & Format(BAT_RS("BRD_DOM"), "DD/MM/YYYY") & "|" & AMT(BAT_RS("PRD_MRP")) & "|" & Trim(INV_RS("STK_TYP")) & "|" & Trim(L_MER_NUM) & "|" & Trim(L_LRN_IDY) & "|" & Format(L_LRN_DAT, "DD/MM/YYYY") & "|" & Trim(INL_RS("TXN_TPT"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIV" & "|" & Val(INV_RS("TOT_PCS")) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIR" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & (Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & (Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) & "|" & Trim(INV_RS("INV_IDY"))
                        K = K + 5
                    End If
                    If Trim(INV_RS("STK_TYP")) = "Free" Then
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIV" & "|" & Val(INV_RS("TOT_PCS")) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SIR" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & (Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & (Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & (Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) & "|" & Trim(INV_RS("INV_IDY"))
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "OMSDA" & "|" & (Val(INV_RS("TOT_PCS")) * Val(L_SAL_LPP)) & "|" & Trim(INV_RS("INV_IDY"))
                        K = K + 6
                        
                        If LNV_RS.State = 1 Then LNV_RS.Close
                        LNV_RS.Open "SELECT * FROM LNV WHERE INV_IDY = '" & Trim(INV_RS("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
                        If LNV_RS.RecordCount > 0 Then
                            While Not LNV_RS.EOF
                                If Val(LNV_RS("LNV_AMT")) > 0 Then
                                    F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "OMSDA" & "|" & Val(LNV_RS("LNV_AMT")) & "|" & Trim(INV_RS("INV_IDY"))
                                    K = K + 1
                                End If
                                LNV_RS.MoveNext
                            Wend
                        End If
                    End If
                End If
                If Trim(INV_RS("TXN_TYP")) = "SRN" Then
                    If Trim(INV_RS("STK_TYP")) = "Salable" Then
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRV" & "|" & (Val(INV_RS("TOT_PCS")) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRR" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        K = K + 5
                    End If
                    If Trim(INV_RS("STK_TYP")) = "Free" Then
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRV" & "|" & (Val(INV_RS("TOT_PCS")) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "SRR" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(BAT_RS("PUR_LPP")) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "RM" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(L_RMM_UNT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "DM" & "|" & ((Val(INV_RS("TOT_PCS")) * (Val(L_DMM_UNT) / Val(BAT_RS("SKU_SBU")))) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "STA" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(PL_TOT_VAT)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        F_INV.WriteLine Format(INV_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & Right(L_CST_IDY, 4) & "|" & Trim(M_WDS_IDY) & "|" & "NA" & "|" & "OMSDA" & "|" & ((Val(INV_RS("TOT_PCS")) * Val(L_SAL_LPP)) * -1) & "|" & "N" & Right(Trim(INV_RS("INV_IDY")), 9)
                        K = K + 6
                    End If
                End If
            End If
            INV_RS.MoveNext
        Wend
'        MsgBox "" & K & ": Stock Records Created for Data Transfer!"
'    Else
'       MsgBox ("No Stock Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
'       cancel
'    End If
    F_INV.Close
    On Error Resume Next
    If INL_RS.State = 1 Then INL_RS.Close
    INL_RS.Open "SELECT * FROM INL WHERE INV_DAT >= CTOD('" & ST_DAT & "') .AND. INV_DAT <= CTOD('" & EN_DAT & "')", con, adOpenKeyset, adLockPessimistic
    If INL_RS.RecordCount > 0 Then INL_RS.MoveFirst
    While Not INL_RS.EOF
        M_ADD_AMT = 0
        M_DIS_AMT = 0
        M_TAX_AMT = 0
        T_SKU_QTY = 0
        If INV_RS.State = 1 Then INV_RS.Close
        INV_RS.Open "SELECT * FROM INV WHERE INV.INV_IDY = '" & Trim(INL_RS("INV_IDY")) & "'", con, adOpenKeyset, adLockPessimistic
        If INV_RS.RecordCount > 0 Then INV_RS.MoveFirst
        While Not INV_RS.EOF
            If INV_RS("PRD_QTY") > 0 Then
                If IsNull(INV_RS("ADD_AMT")) = False Then
                    M_ADD_AMT = M_ADD_AMT + INV_RS("ADD_AMT")
                End If
                If IsNull(INV_RS("DIS_AMT")) = False Then
                    M_DIS_AMT = M_DIS_AMT + INV_RS("DIS_AMT")
                End If
                If IsNull(INV_RS("TAX_AMT")) = False Then
                    M_TAX_AMT = M_TAX_AMT + INV_RS("TAX_AMT")
                End If
                If IsNull(INV_RS("TOT_PCS")) = False Then
                    T_SKU_QTY = Val(T_SKU_QTY) + Val(INV_RS("TOT_PCS"))
                End If
            End If
            INV_RS.MoveNext
        Wend
        If Trim(INL_RS("TXN_TYP")) <> "OPS" Then
            If INV_RS.RecordCount > 0 Then
                If Trim(INL_RS("TXN_TYP")) = "PUR" Then
                    'F_INL.WriteLine Trim(INL_RS("INV_IDY")) & "|" & Format(INL_RS("INV_DAT"), "DD/MM/YYYY") & "|" & INL_RS("TXN_TYP") & "|" & INL_RS("CST_IDY") & "|" & AMT(INL_RS("TAX_AMT") + INL_RS("EXD_AMT") + M_TAX_AMT) & "|" & AMT(INL_RS("SPD_AMT") + INL_RS("CSD_AMT") + M_ADD_AMT + M_DIS_AMT) & "|" & AMT(INL_RS("NET_AMT")) & "|" & Trim(INL_RS("SMN_IDY")) & "|" & Trim(INL_RS("ROU_IDY")) & "|" & Trim(INL_RS("ORD_IDY"))
                    F_INL.WriteLine Format(INL_RS("INV_DAT"), "YYYYMMDD") & "|" & Trim(M_WDS_IDY) & "|" & Trim(INL_RS("INV_IDY")) & "|" & Trim(INL_RS("ORD_IDY")) & "|" & Format(INL_RS("DCH_DAT"), "YYYYMMDD") & "|" & Val(T_SKU_QTY) & "|" & Val(INL_RS("NET_AMT"))
                    r = r + 1
                Else
                    'F_INL.WriteLine Trim(INL_RS("INV_IDY")) & "|" & Format(INL_RS("INV_DAT"), "DD/MM/YYYY") & "|" & INL_RS("TXN_TYP") & "|" & INL_RS("CST_IDY") & "|" & AMT(INL_RS("TAX_AMT") + INL_RS("EXD_AMT") + M_TAX_AMT) & "|" & AMT(INL_RS("SPD_AMT") + INL_RS("CSD_AMT") + M_ADD_AMT + M_DIS_AMT) & "|" & AMT(INL_RS("NET_AMT")) & "|" & Trim(INL_RS("SMN_IDY")) & "|" & Trim(INL_RS("ROU_IDY")) & "|" & ""
                    'r = r + 1
                End If
            End If
        End If
        INL_RS.MoveNext
    Wend
    F_INL.Close
    
'    cancel

    ' * : Beginning of DAS_STOCK.DAT Creation Code for Data Transfer

    Dim F_STK
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CurDir & "\HHT") = False Then
       FS_SOU.CreateFolder (CurDir & "\HHT")
    End If
    
    If FS_SOU.FileExists(CurDir & "\HHT\DAS_STOCK.DAT") = True Then
        FS_SOU.DeleteFile (CurDir & "\HHT\DAS_STOCK.DAT")
    End If
    
    Set F_STK = FS_SOU.OpenTextFile(CurDir & "\HHT\DAS_STOCK.DAT", 2, True)
    Dim l As Integer
    l = 0
    LDD = DateDiff("D", BEG_DAT, END_DAT)
    LDD = LDD + 1
    Do While LDD <> 0
            STKTXTDEX (BEG_DAT)
            Dim CLD_RS As New ADODB.Recordset
            
            If CLD_RS.State = 1 Then CLD_RS.Close
            If PRD_RS.State = 1 Then PRD_RS.Close
            If BAT_RS.State = 1 Then BAT_RS.Close
            CLD_RS.Open "SELECT PRD.PRD_IDY, SUM(CLS_STK) FROM CLD, BRD, PRD WHERE CLD.PRD_IDY = BRD.BAT_IDY AND BRD.PRD_IDY = PRD.PRD_IDY AND CLD.CLS_STK > 0 ORDER BY PRD.PRD_IDY GROUP BY PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
            PRD_RS.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
            BAT_RS.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
            
            
            If CLD_RS.RecordCount > 0 Then CLD_RS.MoveFirst
            
                While Not CLD_RS.EOF
                               
                    If BAT_RS.RecordCount > 0 Then BAT_RS.MoveFirst
                    BAT_RS.Find "BAT_IDY='" & Trim(CLD_RS("PRD_IDY")) & "'"
                    If BAT_RS.EOF = False Then
                       L_BAT_NME = BAT_RS("BAT_NME")
                    End If
                    
                    If PRD_RS.State = 1 Then PRD_RS.Close
                    PRD_RS.Open "SELECT * FROM PRD WHERE PRD.PRD_IDY = '" & Trim(CLD_RS(0)) & "'", con, adOpenKeyset, adLockPessimistic
                    'If PRD_RS.RecordCount > 0 Then PRD_RS.MoveFirst
                    'PRD_RS.Find "PRD_IDY='" & Trim(BAT_RS("PRD_IDY")) & "'"
                    If PRD_RS.EOF = False Then
                       L_PRD_COD = PRD_RS("ALT_IDY")
                    End If
                                
        '            F_STK.WriteLine Trim(M_WDS_IDY) & "|" & Trim(L_PRD_COD) & "|" & Trim(STR(CLD_RS("OPN_STK"))) & "|" & Trim(STR(CLD_RS("PUR_STK"))) & "|" & Trim(STR(CLD_RS("PRT_STK"))) & "|" & Trim(STR(CLD_RS("SAL_STK"))) & "|" & Trim(STR(CLD_RS("SRN_STK"))) & "|" & Trim(STR(CLD_RS("STI_STK"))) & "|" & Trim(STR(CLD_RS("STO_STK"))) & "|" & Trim(STR(CLD_RS("CLS_STK"))) & "|" & Trim(STR(BAT_RS("PRD_MRP"))) & "|" & Trim(BAT_RS("BAT_NME")) & "|" & BAT_RS("BRD_DOM")
                    'F_STK.WriteLine Format(BEG_DAT, "DD/MM/YYYY") & "|" & Trim(L_PRD_COD) & "|" & Trim(L_BAT_NME) & "|" & BAT_RS("BRD_DOM") & "|" & AMT(BAT_RS("PRD_MRP")) & "|" & QTY(STM(CLD_RS("OPN_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("PUR_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("PRT_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("SAL_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("SRN_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("STI_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("STO_STK"), PRD_RS("SUB_UNT"))) & "|" & QTY(STM(CLD_RS("CLS_STK"), PRD_RS("SUB_UNT")))
                    F_STK.WriteLine Format(BEG_DAT, "YYYYMMDD") & "|" & Format(END_DAT, "YYYYMMDD") & "|" & Trim(L_PRD_COD) & "|" & Trim(M_WDS_IDY) & "|" & Val(CLD_RS(1))
                    CLD_RS.MoveNext
                    l = l + 1
                Wend
        '        MsgBox "" & L & ": SKU Records Created for Data Transfer!"
        '    Else
        '       MsgBox ("No Stock Records Available for Data Transfer!"), vbInformation, "DAS Version 4.0"
        '       cancel
        '    End If
           
        '    cancel
            ' * : End of DAS_STOCK.DAT Creation Code for Data Transfer
        BEG_DAT = (BEG_DAT + 1)
        LDD = LDD - 1
    Loop
    F_STK.Close
    
    ' * : Log Text File Creation Begin Process
    
    Dim F_VER
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    If FS_SOU.FileExists(CurDir & "\HHT\DAS_VERSION.DAT") = True Then
        FS_SOU.DeleteFile (CurDir & "\HHT\DAS_VERSION.DAT")
    End If
    
    Dim Q As Integer
    Dim LDC_DAT As Date
    Q = 0
    Set F_VER = FS_SOU.OpenTextFile(CurDir & "\HHT\DAS_VERSION.DAT", 2, True)
    
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * FROM DIR where substr(alltrim(DIR_IDY),1,3) = 'DAY' AND alltrim(DIR_MSG) = 'Y'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveLast
        LDC_DAT = DIR_RS("DIR_DAT")
    Else
        LDC_DAT = Date
    End If
    F_VER.WriteLine Trim(M_WDS_IDY) & "|" & Trim("01.01.01") & "|" & Format(FromDate, "dd/MM/yyyy") & "|" & Format(ToDate, "dd/MM/yyyy") & "|" & Format(LDC_DAT, "dd/MM/yyyy") & "|" & Trim(M_EXE_VER)
    Q = Q + 1
    
    F_VER.Close
    
    Dim F_LOG
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If FS_SOU.FolderExists(CurDir & "\HHT") = False Then
       FS_SOU.CreateFolder (CurDir & "\HHT")
    End If
    
    If FS_SOU.FileExists(CurDir & "\HHT\DAS_LOG.DAT") = True Then
        FS_SOU.DeleteFile (CurDir & "\HHT\DAS_LOG.DAT")
    End If
    
    Set F_LOG = FS_SOU.OpenTextFile(CurDir & "\HHT\DAS_LOG.DAT", 2, True)
    
    F_LOG.WriteLine "DAS_CUSTOMER.DAT" & "|" & Str(J)
    F_LOG.WriteLine "DAS_SECONDARY.DAT" & "|" & Str(K)
    F_LOG.WriteLine "DAS_PURCHASE.DAT" & "|" & Str(r)
    F_LOG.WriteLine "DAS_STOCK.DAT" & "|" & Str(l)
    F_LOG.WriteLine "DAS_VERSION.DAT" & "|" & Str(Q)
    
    Dim MDD As String
    Dim MMM As String
    Dim MYY As String
    MDD = Trim(Str(Day(ToDate)))
    MMM = Trim(Str(Month(ToDate)))
    MYY = Trim(Str(Year(ToDate)))
'    F_LOG.WriteLine MDD & MMM & MYY & M_WDS_SRL 'Refers to Day Closure Date, Last Line in Log File is Blocked - PNR
    F_LOG.Close
    ' * : Log Text File Creation End Process
'    MsgBox ("File Creation Process Successful! Log File Generated!"), vbInformation, "DAS Version 4.0"
    
' Beginning of Distributor Information for MAPI Data Xfer
    
    Dim M_NME, M_ADR, M_TFX, M_WDS As String
    Dim RST_DIR As New ADODB.Recordset
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY = 'SYS' .AND. DIR_IDX = 'TT1'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_NME = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY = 'SYS' .AND. DIR_IDX = 'TT2'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_ADR = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY = 'SYS' .AND. DIR_IDX = 'TT3'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_TFX = Trim(RST_DIR("DIR_MSG"))
    End If
    
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY = 'WDS' .AND. DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        M_WDS = Trim(RST_DIR("DIR_ALS"))
    End If

    If Dir(CurDir & "\DST_INF.DAT") = "" Then
        Set F_VER = FS_SOU.OpenTextFile(CurDir & "\DST_INF.DAT", 2, True)
        F_VER.WriteLine "Distributor : " & M_NME
        F_VER.WriteLine "Address     : " & M_ADR
        F_VER.WriteLine "Telephone   : " & M_TFX
        F_VER.WriteLine "DAS Code-ID : " & M_WDS
        F_VER.Close
    End If

' End of Distributor Information for MAPI Data Xfer

    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open "SELECT * from DIR where DIR_IDY = 'WDS' and DIR_IDX = 'IDY'", con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then DIR_RS.MoveFirst
       DIR_RS("DIR_MSG") = Trim(M_WDS_SRL)
       DIR_RS.Update
    If DIR_RS.State = 1 Then DIR_RS.Close
    'Zip File Creation
            Dim DEX As String
            Dim CDO As String
            DEX = "C:\DEX"
            CDO = "C:\DEX\OutBox"
            If Dir(DEX, vbDirectory) = "" Then
                MkDir DEX
            End If
            If Dir(CDO, vbDirectory) = "" Then
                MkDir CDO
            End If
            
'            CopyFile CURDIR & "\HHT" & "\DAS_CUSTOMER.DAT", CDO & "\DAS_CUSTOMER.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_SECONDARY.DAT", CDO & "\DAS_SECONDARY.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_PURCHASE.DAT", CDO & "\DAS_PURCHASE.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_STOCK.DAT", CDO & "\DAS_STOCK.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_VERSION.DAT", CDO & "\DAS_VERSION.DAT", 0
'            CopyFile CURDIR & "\HHT" & "\DAS_LOG.DAT", CDO & "\DAS_LOG.DAT", 0
'            CopyFile CURDIR & "\DST_INF.DAT", CDO & "\DST_INF.DAT", 0
            
            CopyFile CurDir & "\HHT" & "\DAS_CUSTOMER.DAT", DEX & "\DAS_CUSTOMER.DAT", 0
            CopyFile CurDir & "\HHT" & "\DAS_SECONDARY.DAT", DEX & "\DAS_SECONDARY.DAT", 0
            CopyFile CurDir & "\HHT" & "\DAS_PURCHASE.DAT", DEX & "\DAS_PURCHASE.DAT", 0
            CopyFile CurDir & "\HHT" & "\DAS_STOCK.DAT", DEX & "\DAS_STOCK.DAT", 0
            CopyFile CurDir & "\HHT" & "\DAS_VERSION.DAT", DEX & "\DAS_VERSION.DAT", 0
            CopyFile CurDir & "\HHT" & "\DAS_LOG.DAT", DEX & "\DAS_LOG.DAT", 0
            CopyFile CurDir & "\DST_INF.DAT", DEX & "\DST_INF.DAT", 0
                        
                        
            
            CopyFile App.Path & "\" & "Zip32.dll", CurDir & "\" & "Zip32.dll", 0
            M_EML_IDY = GetValue("DIR_MSG", "DIR where DIR_IDY = 'EML' AND DIR_IDX = 'IDY'")
            
            M_FIL_NME = "\" & "DAS" & Trim(M_WDS_IDY) & Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & ".ZIP"
            Dim Myzip As ColaZip.GSColaZip32
            Set Myzip = New ColaZip.GSColaZip32
            Myzip.RootDirectory = "\"
            
            Myzip.ZipFileName = "C:\DEX" & M_FIL_NME
            Myzip.AddFile ("C:\DEX" & "\DAS_CUSTOMER.DAT")
            Myzip.AddFile ("C:\DEX" & "\DAS_SECONDARY.DAT")
            Myzip.AddFile ("C:\DEX" & "\DAS_PURCHASE.DAT")
            Myzip.AddFile ("C:\DEX" & "\DAS_STOCK.DAT")
            Myzip.AddFile ("C:\DEX" & "\DAS_VERSION.DAT")
            Myzip.AddFile ("C:\DEX" & "\DAS_LOG.DAT")
            Myzip.AddFile ("C:\DEX" & "\DST_INF.DAT")
            
            Myzip.MakeColaZip ("DAS")
            
'            MsgBox ("Zip File " & M_FIL_NME & " Created in the Folder " & DEX), vbInformation, "DAS Version 4.0"
                      
            M_MSG_STR = "Margin Minder Zipped Data From : " & Trim(M_NME) & " , " & Trim(M_ADR) & " , " & Trim(M_TFX) & " , " & Trim(M_WDS) & " , Date: " & Format(Date, ("YYYY/MM/DD")) & " , Time: " & Format(Time, "hh:mm:ss")
            M_MSG_EML = GetValue("DIR_MSG", "DIR WHERE DIR_IDY = 'EML' AND DIR_IDX = 'IDY'")
            M_MSG_SUB = "Zip Data Xfer Margin Minder Data Extract : " & M_WDS & " , Date: " & Format(Date, ("YYYY/MM/DD")) & " , Time: " & Format(Time, "hh:mm:ss")
            
            M_MAIL_ATTMNT = "C:\DEX" & M_FIL_NME
            CopyFile DEX & M_FIL_NME, CDO & M_FIL_NME, 0
    ' Zip File Creation
            
'End If
Unload Me
End Sub

Private Sub STKTXTDEX(DEN_DATE As Date)
    On Error Resume Next
    EN_DAT = DEN_DATE
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    con.Execute "DELETE FROM CLD"
    rs.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockPessimistic
'        RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where inv_dat < ctod('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika and Ramanesh
        RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat < ctod('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic
    For A = 1 To RS1.RecordCount
        If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
            con.Execute "UPDATE CLD SET OPN_STK = OPN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
        con.Execute "UPDATE CLD SET PUR_STK = PUR_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
        con.Execute "UPDATE CLD SET PRT_STK = PRT_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
        con.Execute "UPDATE CLD SET STI_STK = STI_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
        con.Execute "UPDATE CLD SET SRN_STK = SRN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
        ElseIf RS1!TXN_TYP = "STO" Or RS1!TXN_TYP = "ADJ" Then
        con.Execute "UPDATE CLD SET STO_STK = STO_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
    con.Execute "UPDATE CLD SET PUR_STK=0"
    con.Execute "UPDATE CLD SET PRT_STK=0"
    con.Execute "UPDATE CLD SET SAL_STK=0"
    con.Execute "UPDATE CLD SET STO_STK=0"
    con.Execute "UPDATE CLD SET SRN_STK=0"
    con.Execute "UPDATE CLD SET STI_STK=0"
    con.Execute "UPDATE CLD SET CLS_STK=0"
    
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM CLD", con, adOpenKeyset, adLockPessimistic
'        RS1.Open "select sum(PRD_UOM),BAT_IDY,txn_typ,stk_typ from inv where inv_dat >= ctod('" & EN_DAT & "') AND INV_DAT <=CTOD('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika & Ramanesh
        If RS1.State = 1 Then RS1.Close
        RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat >= ctod('" & EN_DAT & "') AND INV_DAT <=CTOD('" & EN_DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Radhika & Ramanesh
    For A = 1 To RS1.RecordCount
        If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
            con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
            con.Execute "UPDATE CLD SET OPN_STK = OPN_STK +" & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
        con.Execute "UPDATE CLD SET PUR_STK = PUR_STK +" & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
        con.Execute "UPDATE CLD SET PRT_STK = PRT_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
        con.Execute "UPDATE CLD SET STI_STK = STI_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
        con.Execute "UPDATE CLD SET SRN_STK = SRN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
        ElseIf RS1!TXN_TYP = "STO" Or RS1!TXN_TYP = "ADJ" Then
        con.Execute "UPDATE CLD SET STO_STK = STO_STK + " & Val(RS1(0)) & " WHERE PRD_IDY = '" & Trim(RS1!BAT_IDY) & "'", d
            If d = 0 Then
                rs.AddNew
                rs!prd_idy = Trim(RS1!BAT_IDY)
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
End Sub

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


Private Sub Form_Load()
    MM_EX_FILE
End Sub
