Attribute VB_Name = "modProcess"
Option Explicit
'Public con As New Connection
Public company As String
Public lblResult As Label
Public lblProcess As Label

Dim xmlString As String
Dim rec As New ADODB.Recordset

Public IVchType(7)  As String
Public VchType(5)   As String

Public TIVchType(7)  As String
Public TVchType(5)   As String

Public Function AddZephyrCompany(cmbZephyrCompany As ComboBox, strFile As String)
    Dim FSO As New FileSystemObject
    Dim txtStream As TextStream
    Dim STR As String
    
    Set txtStream = FSO.OpenTextFile(strFile, ForReading)
    txtStream.ReadLine
    cmbZephyrCompany.CLEAR
    While Not txtStream.AtEndOfLine
        STR = txtStream.ReadLine()
        STR = Mid(STR, 2, InStr(1, STR, "@") - 2)
        cmbZephyrCompany.AddItem STR
    Wend
    txtStream.Close
End Function

Public Function AddTallyCompany(cmbTallyCompany As ComboBox)
    Dim strResponse As String * 255
    strResponse = ModTallyInterface.GetData(modXml.CompanyList)
'--------------if comes in xml format
    Dim XMLDoc As DOMDocument
    Dim nodeList As IXMLDOMNodeList
    
    Dim nCount
    Set XMLDoc = New DOMDocument
    XMLDoc.async = False
    XMLDoc.loadXML (strResponse)
    Set nodeList = XMLDoc.selectNodes("//COMPANYNAME")
    cmbTallyCompany.CLEAR
    If strResponse <> "" And nodeList.length <> 0 Then
        For nCount = 0 To nodeList.length - 1
            cmbTallyCompany.AddItem nodeList.Item(nCount).Text
        Next
    Else
'-----------if comes in text format
        Dim bool As Boolean
        bool = True
        Dim s As String
        Dim I As Integer
        Dim i1 As Integer
        While bool
            i1 = InStr(1, strResponse, Chr(13))
            If (i1 > 0) Then
            s = Mid(strResponse, 1, i1 - 1)
            s = Replace(s, Chr(13), "")
            s = Replace(s, Chr(10), "")
            cmbTallyCompany.AddItem Replace(s, Chr(13), "")
            strResponse = Mid(strResponse, i1 + 1)
            If Len(strResponse) <= 2 Then bool = False
            Else
            bool = False
            End If
        Wend
    End If
End Function



Public Function CreatePrimaryGroups() As Long
'Error Handler
    Dim RESULT As Long
    RESULT = 0
    On Error GoTo ErrHandler
    DoEvents
    
    Dim FSO As FileSystemObject
    Dim txtStream As TextStream
    Dim strLine As String
    Dim LedgerName As String
    Dim IsDeemedPostivie As String
    Dim isRevenue As String
    Dim temploc As Integer
    Dim temp As String
    
    
    Set FSO = New FileSystemObject
    Set txtStream = FSO.OpenTextFile(SetupFolder & "\PrimaryGroups.txt", ForReading)
    strLine = txtStream.ReadLine
    strLine = txtStream.ReadLine
    strLine = txtStream.ReadLine
    strLine = txtStream.ReadLine
    strLine = txtStream.ReadLine
    strLine = txtStream.ReadLine
    Do While Not txtStream.AtEndOfLine
        strLine = txtStream.ReadLine
        
        temploc = InStr(1, strLine, ":")
        temp = Mid(strLine, 1, temploc - 1)
        LedgerName = temp
        strLine = Mid(strLine, temploc + 1)
        
        temploc = InStr(1, strLine, ":")
        temp = Mid(strLine, 1, temploc - 1)
        IsDeemedPostivie = temp
        strLine = Mid(strLine, temploc + 1)
        
        isRevenue = strLine
        DisplayResult "         Primary Account Group - " & LedgerName
        
        xmlString = ""
        xmlString = modXml.AddOpeningTags(company, "All Masters") & xmlString
        If InStr(1, UCase(Trim(LedgerName)), "CASH") > 0 Then
            xmlString = xmlString & modXml.ConvertGroupToXML("Create", LedgerName, LedgerName, "Cash-in-hand")
        ElseIf InStr(1, UCase(Trim(LedgerName)), "BANK") > 0 Then
            xmlString = xmlString & modXml.ConvertGroupToXML("Create", LedgerName, LedgerName, "Bank Accounts")
        ElseIf InStr(1, UCase(Trim(LedgerName)), "PURCHASES") > 0 Then
            xmlString = xmlString & modXml.ConvertGroupToXML("Create", LedgerName, LedgerName, "Purchase Accounts")
        ElseIf InStr(1, UCase(Trim(LedgerName)), "SALES") > 0 Then
            xmlString = xmlString & modXml.ConvertGroupToXML("Create", LedgerName, LedgerName, "Sales Accounts")
        ElseIf InStr(1, UCase(Trim(LedgerName)), "CREDITOR") > 0 Then
            xmlString = xmlString & modXml.ConvertGroupToXML("Create", LedgerName, LedgerName, "Sundry Creditors")
        ElseIf InStr(1, UCase(Trim(LedgerName)), "DEBTOR") > 0 Then
            xmlString = xmlString & modXml.ConvertGroupToXML("Create", LedgerName, LedgerName, "Sundry Debtors")
        ElseIf InStr(1, UCase(Trim(LedgerName)), "BANK OVERDRAFT") > 0 Then
            xmlString = xmlString & modXml.ConvertGroupToXML("Create", LedgerName, LedgerName, "Band OD A/c")
        ElseIf InStr(1, UCase(Trim(LedgerName)), "ASSET") > 0 Then
            xmlString = xmlString & modXml.ConvertGroupToXML("Create", LedgerName, LedgerName, "Fixed Assets")
        ElseIf InStr(1, UCase(Trim(LedgerName)), "LIABILITY") > 0 Then
            xmlString = xmlString & modXml.ConvertGroupToXML("Create", LedgerName, LedgerName, "Loans (Liability)")
        Else
            xmlString = xmlString & modXml.ConvertGroupToXML("Create", LedgerName, LedgerName, "", IsDeemedPostivie, isRevenue)
        End If
        xmlString = xmlString & modXml.AddClosingTags()
        RESULT = ModTallyInterface.SendData(xmlString)
    Loop
    CreatePrimaryGroups = RESULT
    Exit Function
ErrHandler:
    CreatePrimaryGroups = RESULT
End Function

Public Function CreateGroups(strParameterValue As String) As Long
If strParameterValue = "" Then
    CreateGroups = 0
    Exit Function
End If
'Error Handler
'    On Error Resume Next
    Dim RESULT As Long
    ReSetConnectionObjects EFF_LIB1.CompanyName
    If rec.State = 1 Then rec.Close
    rec.Open "select  DISTINCT grp_idy, grp_nme, grp_clf from grp where grp_idy in (" & strParameterValue & ")  order by grp_idy", CON, adOpenForwardOnly, adLockReadOnly
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            DisplayResult "         Account Group - " & rec("grp_nme")
            xmlString = ""
            If UCase(Trim(rec("grp_nme"))) = "BANK ACCOUNTS" Or UCase(Trim(rec("grp_nme"))) = "BANK OCC A/C" Or UCase(Trim(rec("grp_nme"))) = "BANK OD A/C" Or UCase(Trim(rec("grp_nme"))) = "BRANCH / DIVISIONS" Or _
                UCase(Trim(rec("grp_nme"))) = "CAPITAL ACCOUNT" Or UCase(Trim(rec("grp_nme"))) = "CASH-IN-HAND" Or UCase(Trim(rec("grp_nme"))) = "CURRENT ASSETS" Or _
                UCase(Trim(rec("grp_nme"))) = "CURRENT LIABILITIES" Or UCase(Trim(rec("grp_nme"))) = "DEPOSITS (ASSET)" Or UCase(Trim(rec("grp_nme"))) = "DIRECT EXPENSES" Or _
                UCase(Trim(rec("grp_nme"))) = "DIRECT INCOMES" Or UCase(Trim(rec("grp_nme"))) = "DUTIES & TAXES" Or UCase(Trim(rec("grp_nme"))) = "EXPENSES (DIRECT)" Or _
                UCase(Trim(rec("grp_nme"))) = "EXPENSES (INDIRECT)" Or UCase(Trim(rec("grp_nme"))) = "FIXED ASSETS" Or UCase(Trim(rec("grp_nme"))) = "INCOME (DIRECT)" Or _
                UCase(Trim(rec("grp_nme"))) = "INCOME (INDIRECT)" Or UCase(Trim(rec("grp_nme"))) = "INDIRECT EXPENSES" Or UCase(Trim(rec("grp_nme"))) = "INDIRECT INCOMES" Or _
                UCase(Trim(rec("grp_nme"))) = "INVESTMENTS" Or UCase(Trim(rec("grp_nme"))) = "LOANS & ADVANCES (ASSET)" Or UCase(Trim(rec("grp_nme"))) = "LOANS (LIABILITY)" Or _
                UCase(Trim(rec("grp_nme"))) = "MISC. EXPENSES (ASSET)" Or UCase(Trim(rec("grp_nme"))) = "PROVISIONS" Or UCase(Trim(rec("grp_nme"))) = "PURCHASE ACCOUNTS" Or _
                UCase(Trim(rec("grp_nme"))) = "RESERVES & SURPLUS" Or UCase(Trim(rec("grp_nme"))) = "RETAINED EARNINGS" Or UCase(Trim(rec("grp_nme"))) = "SALES ACCOUNTS" Or _
                UCase(Trim(rec("grp_nme"))) = "SECURED LOANS" Or UCase(Trim(rec("grp_nme"))) = "STOCK-IN-HAND" Or UCase(Trim(rec("grp_nme"))) = "SUNDRY CREDITORS" Or _
                UCase(Trim(rec("grp_nme"))) = "SUNDRY DEBTORS" Or UCase(Trim(rec("grp_nme"))) = "SUSPENSE A/C" Or UCase(Trim(rec("grp_nme"))) = "UNSECURED LOANS" Or _
                UCase(Trim(rec("grp_nme"))) = "INCOME INDIRECT" Or UCase(Trim(rec("grp_nme"))) = "EXPENSES DIRECT" Or UCase(Trim(rec("grp_nme"))) = "INCOME DIRECT" Or _
                UCase(Trim(rec("grp_nme"))) = "EXPENSES INDIRECT" Or UCase(Trim(rec("grp_nme"))) = "LOANS LIABILITY" Or UCase(Trim(rec("grp_nme"))) = "LOANS & ADVANCES ASSET" Then
            Else
                xmlString = modXml.AddOpeningTags(company, "All Masters") & xmlString
'                If InStr(1, UCase(Trim(rec("grp_nme"))), "CASH") > 0 Then
'                    xmlString = xmlString & modXml.ConvertGroupToXML("Create", UCase(Trim(rec("grp_nme"))), UCase(Trim(rec("grp_nme"))), "Cash-in-hand")
'                ElseIf InStr(1, UCase(Trim(rec("grp_nme"))), "BANK") > 0 Then
'                    xmlString = xmlString & modXml.ConvertGroupToXML("Create", UCase(Trim(rec("grp_nme"))), UCase(Trim(rec("grp_nme"))), "Bank Accounts")
'                Else
                    xmlString = xmlString & modXml.ConvertGroupToXML("Create", UCase(Trim(rec("grp_nme"))), UCase(Trim(rec("grp_nme"))), Trim(rec("grp_clf")))
'                End If
                'xmlstring = xmlstring & modXml.ConvertGroupToXML("Create", rec("grp_nme"), rec("grp_nme"), Replace(TrimSpaces(rec("grp_clf")), "&", "-"))
                xmlString = xmlString & modXml.AddClosingTags()
                RESULT = ModTallyInterface.SendData(xmlString)
            End If
            rec.MoveNext
        Loop
    End If
    If rec.State = 1 Then rec.Close
    CreateGroups = RESULT
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If RESULT <> 0 Then RESULT = err.Number
    CreateGroups = RESULT
End Function

Public Function CreateAccounts(strParameterValue As String) As Long
'Error Handler
    On Error GoTo ErrHandler
    Dim RESULT As Long
    Dim rec1 As New ADODB.Recordset
    Dim amt As Double
    
    RESULT = 0
    rec.Open "select  DISTINCT act_idy, act_nme, grp_nme from act,grp where grp.grp_idy = act.grp_idy and act_idy in (" & strParameterValue & ")  order by act_idy", CON, adOpenForwardOnly, adLockReadOnly
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            amt = 0
            'OPN - Opening Balance
            If rec1.State = 1 Then rec1.Close
            rec1.Open "select sum(txn_amt) as opn from vch where crt_idy = '" & rec("act_idy") & "' and txn_typ='OPN'", CON, adOpenStatic, adLockOptimistic
            If Not rec1.EOF Then amt = IIf(IsNull(rec1("opn")) = True, 0, rec1("opn"))
            rec1.Close
            If amt = 0 Then
                rec1.Open "select sum(txn_amt) as opn from vch where dbt_idy = '" & rec("act_idy") & "' and txn_typ='OPN'", CON, adOpenStatic, adLockOptimistic
                If Not rec1.EOF Then amt = IIf(IsNull(rec1("opn")) = True, 0, rec1("opn"))
                rec1.Close
            End If
            'OPS - Opening Stock Balance
            If amt = 0 Then
                rec1.Open "select sum(net_amt) as opn from inl where act_idy = '" & rec("act_idy") & "' and txn_typ='OPS'", CON, adOpenForwardOnly, adLockReadOnly
                If Not rec1.EOF Then amt = IIf(IsNull(rec1("opn")) = True, 0, rec1("opn"))
                rec1.Close
            End If
            If amt = 0 Then
                rec1.Open "select sum(net_amt) as opn from inl where cst_idy = '" & rec("act_idy") & "' and txn_typ='OPS'", CON, adOpenForwardOnly, adLockReadOnly
                If Not rec1.EOF Then amt = IIf(IsNull(rec1("opn")) = True, 0, rec1("opn"))
                rec1.Close
            End If
            If rec1.State = 1 Then rec1.Close
            RESULT = CreateAccount("Create", Trim(rec("act_nme")), Trim(rec("act_nme")), Trim(rec("grp_nme")), amt)
'            result = CreateAccount("Create", Trim(rec("act_nme")), Trim(rec("act_nme")), Trim(rec("grp_nme")), amt)
            rec.MoveNext
        Loop
    End If
    RESULT = CreateAccount("Create", "VAT Output", "VAT Output", "Duties & Taxes", 0)
    RESULT = CreateAccount("Create", "VAT Input", "VAT Input", "Duties & Taxes", 0)
    RESULT = CreateAccount("Create", "Output VAT @ 4%", "Output VAT @  4%", "Duties & Taxes", 0)
    RESULT = CreateAccount("Create", "Output VAT @ 12.5%", "Output VAT @ 12.5%", "Duties & Taxes", 0)
    RESULT = CreateAccount("Create", "Output VAT @ 10%", "Output VAT @ 10%", "Duties & Taxes", 0)
    RESULT = CreateAccount("Create", "Output VAT @ 2%", "Output VAT @ 2%", "Duties & Taxes", 0)
    RESULT = CreateAccount("Create", "Input VAT @ 4%", "Input VAT @ 4%", "Duties & Taxes", 0)
    RESULT = CreateAccount("Create", "Input VAT @ 12.5%", "Input VAT @ 12.5%", "Duties & Taxes", 0)
    RESULT = CreateAccount("Create", "Input VAT @ 10%", "Input VAT @ 10%", "Duties & Taxes", 0)
    RESULT = CreateAccount("Create", "Input VAT @ 2%", "Input VAT @ 2%", "Duties & Taxes", 0)
    RESULT = CreateAccount("Create", "Direct Expenses", "Direct Expenses", "Direct Expenses", 0)
    RESULT = CreateAccount("Create", "Direct Incomes", "Direct Incomes", "Direct Incomes", 0)
    If rec.State = 1 Then rec.Close
    CreateAccounts = RESULT
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If RESULT <> 0 Then RESULT = err.Number
    CreateAccounts = RESULT
End Function

Public Function CreateUOM(strParameterValue As String) As Long
'Error Handler
    On Error GoTo ErrHandler
    Dim RESULT As Long
    RESULT = 0
    rec.Open "SELECT DISTINCT  PRD_UOM FROM PRD WHERE TRIM(PRD_UOM) IN (" & strParameterValue & ") ORDER BY PRD_UOM ", CON, adOpenForwardOnly, adLockReadOnly
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            DisplayResult "         UOM - " & rec("prd_uom")
            xmlString = ""
            xmlString = modXml.AddOpeningTags(company, "All Masters") & xmlString
            xmlString = xmlString & modXml.ConvertUnitToXML("Create", Replace(rec("prd_uom"), "&", "-"))
            xmlString = xmlString & modXml.AddClosingTags()
            RESULT = ModTallyInterface.SendData(xmlString)
            rec.MoveNext
        Loop
    End If
    If rec.State = 1 Then rec.Close
    CreateUOM = RESULT
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If RESULT <> 0 Then RESULT = err.Number
    CreateUOM = RESULT
End Function

Public Function CreateStockGroup(strParameterValue As String) As Long
'Error Handler
    On Error GoTo ErrHandler
    Dim RESULT As Long
    RESULT = 0
    rec.Open "select  DISTINCT  grp_cls from pgr where grp_idy in (" & strParameterValue & ") order by grp_cls", CON, adOpenForwardOnly, adLockReadOnly
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            DisplayResult "         Stock Group - " & rec("grp_cls")
            xmlString = ""
            xmlString = modXml.AddOpeningTags(company, "All Masters") & xmlString
            xmlString = xmlString & modXml.ConvertStockGroupToXML("Create", Replace(rec("grp_cls"), "&", "-"), Replace(rec("grp_cls"), "&", "-"))
            xmlString = xmlString & modXml.AddClosingTags()
            RESULT = ModTallyInterface.SendData(xmlString)
            rec.MoveNext
        Loop
    End If
    If rec.State = 1 Then rec.Close
    CreateStockGroup = RESULT
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If RESULT <> 0 Then RESULT = err.Number
    CreateStockGroup = RESULT
End Function

Public Function CreateStockItem(strParameterValue As String) As Long
'Error Handler
    On Error GoTo ErrHandler
    Dim OpeningBalance As Double
    Dim OpeningValue As Double
    
    Dim RESULT As Long
    RESULT = 0
    Dim rec1 As New Recordset
    If rec.State = 1 Then rec.Close
    rec.Open "select  DISTINCT  prd_idy, prd_nme, PGR.GRP_CLS, prd_uom, pur_rat, sal_rat, mrp_rat, prd.cur_lev, prd.reo_lev, prd.saf_lev, prd.PUR_DIS , prd.SAL_DIS from pgr, prd where pgr.grp_idy = prd.grp_idy and PRD.PRD_IDY in (" & strParameterValue & ")  order by prd_nme", CON, adOpenKeyset, adLockPessimistic
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            DisplayResult "         Stock Item (SKU) - " & rec("prd_nme")
            xmlString = ""
            xmlString = modXml.AddOpeningTags(company, "All Masters") & xmlString
            Dim reo As Double
            If IsNull(rec("reo_lev")) = True Then
                reo = 0
            Else
                reo = CDbl(rec("reo_lev"))
            End If
            If rec1.State = 1 Then rec1.Close
            rec1.Open "SELECT SUM(PRD_QTY) AS PRD_QTY, SUM(COL_AMT) AS COL_AMT FROM INV,PRD  WHERE  TRIM(INV.TXN_TYP)= 'OPS' AND INV.PRD_IDY = PRD.PRD_IDY AND TRIM(PRD.PRD_IDY) = '" & Trim(rec!prd_idy) & "'", CON, adOpenKeyset, adLockPessimistic
                OpeningBalance = 0
                OpeningValue = 0
            If rec1.EOF = False Then
                If IsNull(rec1!prd_qty) = False Then
                    OpeningBalance = CDbl(rec1!prd_qty)
                    OpeningValue = -CDbl(rec1!COL_AMT)
                End If
            End If
            xmlString = xmlString & modXml.ConvertStockItemToXML("Create", Replace(rec("prd_nme"), "&", "-"), Replace(rec("prd_nme"), "&", "-"), Replace(rec("grp_cls"), "&", "-"), Replace(rec("prd_uom"), "&", "-"), OpeningBalance, reo, CDbl(rec("saf_lev")), CDbl(rec("pur_rat")), CDbl(rec("sal_rat")), OpeningValue)
            xmlString = xmlString & modXml.AddClosingTags()
            RESULT = ModTallyInterface.SendData(xmlString)
            rec.MoveNext
        Loop
    End If
    If rec.State = 1 Then rec.Close
    CreateStockItem = RESULT
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If RESULT <> 0 Then RESULT = err.Number
    CreateStockItem = RESULT
End Function

Public Function CreateAccountingVoucher(VoucherType As String, strParameterValue As String) As Long
'Tally Voucher Types
'JRN - Journal
'PMT - Payment
'RCT - Receipt
'Error Handler


    Dim TallyVchType As String
    Dim xmlstring1 As String
    Dim sql1 As String
    Dim RESULT As Long
    Dim DrAmount As String
    Dim CrAmount As String
    Dim DrAccount As String
    Dim CrAccount As String
    Dim PartyLedgerName As String
    Dim IsDeemedPositive As String
    Dim aid As String
    
    RESULT = 0
    If VoucherType = "OPN" Then Exit Function
    Select Case VoucherType
    Case "PMT"
        TallyVchType = "Payment"
    Case "RCT"
        TallyVchType = "Receipt"
    Case "JRN"
        TallyVchType = "Journal"
    Case "CRN"
        TallyVchType = "Credit Note"
    Case "DBN"
        TallyVchType = "Debit Note"
    End Select

    On Error GoTo ErrHandler
    sql1 = ""
    Dim I As Integer
    
    sql1 = "select txn_idy, txn_dat, txn_typ, txn_amt, a1.act_nme as dr_nme, a2.act_nme as cr_nme, " & _
            " a2.act_idy as caid, a1.act_idy as daid  " & _
            " from vch, act  a1, act  a2 " & _
            " where vch.dbt_idy = a1.act_idy And vch.crt_idy = a2.act_idy  " & _
            " and  VCH.TXN_IDY  in (" & strParameterValue & ")" & _
            " and vch.txn_typ='" & VoucherType & _
            "' order by txn_idy"
    rec.Open sql1, CON, adOpenKeyset, adLockPessimistic
    If Not rec.EOF Then
        Do While Not rec.EOF
            I = I + 1
            DoEvents
            xmlString = ""
            DrAmount = "-0.00"
            CrAmount = "0.00"
            DrAmount = CStr(-1 * CDbl(rec("txn_amt")))
            CrAmount = CStr(CDbl(rec("txn_amt")))
            DrAccount = ""
            CrAccount = ""
            CrAccount = Trim(rec("Cr_nme"))
            DrAccount = Trim(rec("Dr_nme"))
            
            CrAccount = Replace(CrAccount, "&", "&amp;")
            CrAccount = Replace(CrAccount, "'", "&apos;")
            CrAccount = Replace(CrAccount, "'", "&lt;")
            CrAccount = Replace(CrAccount, "'", "&gt;")
            CrAccount = Replace(CrAccount, "'", "&quot;")
           
            DrAccount = Replace(DrAccount, "&", "&amp;")
            DrAccount = Replace(DrAccount, "'", "&apos;")
            DrAccount = Replace(DrAccount, "'", "&lt;")
            DrAccount = Replace(DrAccount, "'", "&gt;")
            DrAccount = Replace(DrAccount, "'", "&quot;")


            PartyLedgerName = CrAccount
            If VoucherType = "RCT" Or VoucherType = "DBN" Then
                aid = Trim(rec("caid"))
            End If
            If VoucherType = "PMT" Or VoucherType = "CRN" Then
                aid = Trim(rec("daid"))
            End If
            If VoucherType = "JRN" Then
                PartyLedgerName = ""
            End If
            DisplayResult "         Voucher Type - " & TallyVchType & "  ID - " & Trim(rec("txn_idy"))
            xmlString = xmlString & modXml.AddOpeningTags(company, "All Masters", TallyVchType)
            '"58392016-768y-675f-867i-597825398422-" &
            xmlString = xmlString & modXml.VoucherOpeningTag("Create", TallyVchType, Trim(rec("txn_idy")) & "-" & aid & " - " & I, Trim(rec("txn_idy")), Trim(PartyLedgerName), Trim(rec("txn_dat")))
            If VoucherType = "RCT" Or VoucherType = "DBN" Then
                aid = Trim(rec("caid"))
                xmlString = xmlString & modXml.ConvertVoucherToXML(CrAccount, CrAmount, "No")
                xmlString = xmlString & modXml.ConvertVoucherToXML(DrAccount, DrAmount, "Yes")
            ElseIf VoucherType = "PMT" Or VoucherType = "CRN" Then
                aid = Trim(rec("daid"))
                xmlString = xmlString & modXml.ConvertVoucherToXML(DrAccount, DrAmount, "Yes")
                xmlString = xmlString & modXml.ConvertVoucherToXML(CrAccount, CrAmount, "No")
            Else
                aid = Trim(rec("caid"))
                xmlString = xmlString & modXml.ConvertVoucherToXML(CrAccount, CrAmount, "No")
                xmlString = xmlString & modXml.ConvertVoucherToXML(DrAccount, DrAmount, "Yes")
            End If
            
            xmlString = xmlString & modXml.VoucherClosingTag()
            xmlString = xmlString & modXml.AddClosingTags()
            RESULT = ModTallyInterface.SendData(xmlString)
            rec.MoveNext
        Loop
    End If
    '------------------------------------
    If rec.State = 1 Then rec.Close
    CreateAccountingVoucher = RESULT
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If RESULT <> 0 Then RESULT = err.Number
    CreateAccountingVoucher = RESULT
End Function


'Inventory vouchers
Public Function CreateInventoryVoucher(VoucherType As String, strParameterValue As String) As Long
'Tally Voucher Types
'PUR - Purchase
'SAL - Sales
'SRN - Credit Note
'PRT - Debit Note
'STI -
'STO -

    If VoucherType = "OPS" Then Exit Function
    Dim TallyVchType As String
    Dim xmlstring1 As String
    Dim rec1 As New ADODB.Recordset
    Dim cramt As Double
    Dim dramt As Double

    Dim sql1 As String
    Dim sql2 As String
    Dim RESULT As Long
    Dim PartyLedgerName As String
    Dim aid As String
    Dim IsDeemedPositive As String
    
    Dim AC1 As String, AC2 As String
    Dim PrdValAmt As Double, PrdDisAmt As Double, PrdTaxAmt As Double
    
    Dim CsdSpdAmt As Double, InlTaxAmt As Double, ExdAmt As Double, PkgCstAmt As Double
    Dim FrtAmt As Double, RndOffAmt As Double, NetAmt As Double
    
    Dim DisB As String, InvTax As String
    Dim DisA As String, InlTax As String, Exd As String, TxnRm1 As String
    Dim TxnRm2 As String, RndOff As String
    Dim AddDis As Double
    
    Dim T As Double
    Dim strTaxType As String
    strTaxType = "Output"
    RESULT = 0

    Select Case VoucherType
    Case "PUR"
        strTaxType = "Input"
        TallyVchType = "Purchase"
    Case "SAL"
        strTaxType = "Output"
        TallyVchType = "Sales"
    Case "SRN"
        TallyVchType = "Credit Note"
    Case "PRT"
        TallyVchType = "Debit Note"
    Case "STI"
        TallyVchType = "Contra"
    Case "STO"
        TallyVchType = "Contra"
    End Select
    
    
    
           
    On Error GoTo ErrHandler
    sql1 = ""
    sql1 = "select INL.inv_idy, INL.INV_DAT, INL.txn_typ, sum(inv.col_amt) as prd_val, sum(INV.DIS_AMT) as prd_dis, SUM(INV.ADD_AMT) as add_dis, " & _
            " a2.act_idy as caid, a1.act_idy as daid,  " & _
            " sum(inv.tax_amt) as prd_tax,  INL.grs_amt, INL.spd_pcg, INL.spd_amt, INL.csd_pcg, INL.csd_amt, INL.exd_pcg, " & _
            " INL.exd_amt, INL.tax_pcg, INL.tax_amt as inl_tax_amt, INL.net_amt, INL.frt_amt, INL.pkg_cst, INL.rou_off , INL.cst_idy, INL.act_idy, a1.act_nme as AC1, a2.act_nme as AC2," & _
            " 'Tax Output @ '+ trim(SUBSTR(str(INL.exd_pcg),5,2)) + '%' + 'A/c'  as exd, " & _
            " 'VAT " & strTaxType & "' as invtax, 'Tax Output @ '+ trim(SUBSTR(str(INL.tax_pcg),5,2)) + '%' + 'A/c'  as inltax, trim(txn_rm1) + ' A/c'  as txnrm1," & _
            " trim(txn_rm2) + ' A/c' as txnrm2, 'Discount A/c (A)' as disA, 'Discount A/c (B)' as disB, " & _
            " 'Round Off A/c'  as roundoff  " & _
            " from inl, inv, act  a1, act  a2 " & _
            " where INL.inv_idy = inv.inv_idy and INL.cst_idy = a1.act_idy and INL.act_idy = a2.act_idy " & _
            " and INL.txn_typ='" & VoucherType & "' and INL.inv_idy in (" & strParameterValue & ") "
    sql1 = sql1 & " group by INL.inv_idy, INL.INV_DAT, INL.txn_typ, INL.cst_idy, INL.act_idy, INL.grs_amt, INL.spd_pcg,  INL.spd_amt, INL.csd_pcg, INL.csd_amt,  INL.exd_pcg, INL.exd_amt, INL.tax_pcg, " & _
            " INL.tax_amt,  INL.net_amt, INL.frt_amt, INL.pkg_cst,INL.rou_off,a1.act_nme, a2.act_nme, txn_rm1, txn_rm2 , a1.act_idy, a2.act_idy"
        If rec.State = 1 Then rec.Close
        rec.Open sql1, CON, adOpenKeyset, adLockPessimistic
    
    Dim I As Integer
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            xmlString = ""
            I = I + 1
        
        
        AC1 = Trim(rec("AC1"))
        AC2 = Trim(rec("AC2"))
        
        AC1 = Replace(AC1, "&", "&amp;")
        AC1 = Replace(AC1, "'", "&apos;")
        AC1 = Replace(AC1, "'", "&lt;")
        AC1 = Replace(AC1, "'", "&gt;")
        AC1 = Replace(AC1, "'", "&quot;")
        
        AC2 = Replace(AC2, "&", "&amp;")
        AC2 = Replace(AC2, "'", "&apos;")
        AC2 = Replace(AC2, "'", "&lt;")
        AC2 = Replace(AC2, "'", "&gt;")
        AC2 = Replace(AC2, "'", "&quot;")

        
        PrdValAmt = rec("prd_val")
        PrdDisAmt = rec("prd_dis") '+ rec("add_dis")
        CsdSpdAmt = rec("csd_amt") + rec("spd_amt")
        PrdValAmt = PrdValAmt - PrdDisAmt - CsdSpdAmt
        PrdTaxAmt = rec("prd_tax")
        
        DisB = rec("disB")
        InvTax = rec("invtax")
        DisA = rec("disA")
        InlTax = rec("inltax")
        Exd = rec("exd")
        TxnRm1 = rec("txnrm1")
        TxnRm2 = rec("txnrm2")
        RndOff = rec("roundoff")
        
        InlTaxAmt = rec("inl_tax_amt")
        ExdAmt = rec("exd_amt")
        PkgCstAmt = rec("pkg_cst")
        FrtAmt = rec("frt_amt")
        RndOffAmt = rec("rou_off")
        NetAmt = rec("net_amt")
        If NetAmt > 0 Then
'------------------------------------------------------------------------------
'------------------------------  ledger Creation  -----------------------------
'------------------------------------------------------------------------------
            RESULT = CreateAccount("Create", DisB, DisB, "Expenses(Indirect)", 0)
            If RESULT <> 0 Then GoTo ErrHandler
            If (PrdTaxAmt > 0) Then RESULT = CreateAccount("Create", InvTax, InvTax, "Duties & Taxes", 0)
            RESULT = CreateAccount("Create", DisA, DisA, "Expenses(Indirect)", 0)
            If InlTaxAmt > 0 Then RESULT = CreateAccount("Create", InlTax, InlTax, "Duties & Taxes", 0)
            If ExdAmt > 0 Then RESULT = CreateAccount("Create", Exd, Exd, "Duties & Taxes", 0)
            If PkgCstAmt > 0 Then RESULT = CreateAccount("Create", TxnRm1, TxnRm1, "Duties & Taxes", 0)
            If FrtAmt > 0 Then RESULT = CreateAccount("Create", TxnRm2, TxnRm2, "Duties & Taxes", 0)
            RESULT = CreateAccount("Create", RndOff, RndOff, "Expenses(Indirect)", 0)
'------------------------------------------------------------------------------
'------------------------------  Ledger Creation  -----------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------  Voucher Creation -----------------------------
'------------------------------------------------------------------------------
            PrdDisAmt = PrdDisAmt
            CsdSpdAmt = CsdSpdAmt
            FrtAmt = FrtAmt
            PrdTaxAmt = PrdTaxAmt
            InlTaxAmt = InlTaxAmt
            PrdValAmt = PrdValAmt
            
            NetAmt = -NetAmt
            ExdAmt = -ExdAmt
            PkgCstAmt = -PkgCstAmt
            RndOffAmt = -(PrdValAmt + NetAmt + FrtAmt + PrdTaxAmt + InlTaxAmt + ExdAmt + PkgCstAmt) ' RndOffAmt
           
            
            If VoucherType = "PUR" Or VoucherType = "PRT" Then
                PrdDisAmt = -PrdDisAmt
                CsdSpdAmt = -CsdSpdAmt
                FrtAmt = -FrtAmt
                PrdTaxAmt = -PrdTaxAmt
                PrdValAmt = -PrdValAmt
                InlTaxAmt = -InlTaxAmt

                NetAmt = -NetAmt
                ExdAmt = -ExdAmt
                PkgCstAmt = -PkgCstAmt
                RndOffAmt = -RndOffAmt
    
                PartyLedgerName = AC2
                aid = Trim(rec("caid"))
            End If
            If VoucherType = "SAL" Or VoucherType = "SRN" Then
    '            PrdDisAmt = -PrdDisAmt
    '            CsdSpdAmt = -CsdSpdAmt
    '            FrtAmt = -FrtAmt
    '            PrdTaxAmt = -PrdTaxAmt
    '            PrdValAmt = -PrdValAmt
    '            InlTaxAmt = -InlTaxAmt
    '
    '            NetAmt = -NetAmt
    '            ExdAmt = -ExdAmt
    '            PkgCstAmt = -PkgCstAmt
    '            RndOffAmt = -RndOffAmt
                PartyLedgerName = Trim(AC1)
                aid = Trim(rec("daid"))
               
            End If
            DisplayResult "         Voucher Type - " & TallyVchType & "  ID - " & Trim(rec("inv_idy"))
            xmlString = ""
            xmlString = xmlString & modXml.AddOpeningTags(company, "All Masters", TallyVchType)
            '"6593251k-803u-731g-953p-597825398422-" &
            xmlString = xmlString & modXml.VoucherOpeningTag("Create", TallyVchType, Trim(rec("inv_idy")) & "-" & aid, Trim(rec("inv_idy")), Trim(PartyLedgerName), Trim(rec("inv_dat")))
            
            xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(AC1), Format(NetAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(AC2), Format(PrdValAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "Yes", "No"))
            If PrdTaxAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(InvTax), Format(PrdTaxAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            If InlTaxAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(InlTax), Format(InlTaxAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            If ExdAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(Exd), Format(ExdAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            If PkgCstAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(TxnRm1), Format(PkgCstAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            If FrtAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(TxnRm2), Format(FrtAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            If RndOffAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(RndOff), Format(RndOffAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
'            If CsdSpdAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(DisA), Format(CsdSpdAmt, "############.00"), IIf(CsdSpdAmt <= 0, "Yes", "No"))
'            If PrdDisAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(DisB), Format(PrdDisAmt, "############.00"), IIf(PrdDisAmt <= 0, "Yes", "No"))
            
            xmlString = xmlString & modXml.VoucherClosingTag()
            xmlString = xmlString & modXml.AddClosingTags()
        
            If modGeneral.strMsg = "1" Then
                Dim FSO As New FileSystemObject
                Dim ts As TextStream
                Set ts = FSO.OpenTextFile(SetupFolder & "\" & Trim(rec("inv_idy")) & ".xml", ForAppending, True)
                ts.Write xmlString
                ts.Close
            End If
            RESULT = ModTallyInterface.SendData(xmlString)
        End If
        rec.MoveNext
    Loop
End If
    '------------------------------------
        If rec.State = 1 Then rec.Close
        CreateInventoryVoucher = RESULT
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If RESULT <> 0 Then RESULT = err.Number
    CreateInventoryVoucher = RESULT
End Function


Private Function CreateAccount(TransactionType As String, LedgerName As String, OriginalLedgerName As String, ParentGroupName As String, OpeningBalance As Double) As Long
    Dim xmlString As String
    DoEvents
    DisplayResult "         Account - " & LedgerName
    xmlString = ""
    
    LedgerName = Replace(LedgerName, "&", "&amp;")
    LedgerName = Replace(LedgerName, "'", "&apos;")
    LedgerName = Replace(LedgerName, "'", "&lt;")
    LedgerName = Replace(LedgerName, "'", "&gt;")
    LedgerName = Replace(LedgerName, "'", "&quot;")

    OriginalLedgerName = Replace(OriginalLedgerName, "&", "&amp;")
    OriginalLedgerName = Replace(OriginalLedgerName, "'", "&apos;")
    OriginalLedgerName = Replace(OriginalLedgerName, "'", "&lt;")
    OriginalLedgerName = Replace(OriginalLedgerName, "'", "&gt;")
    OriginalLedgerName = Replace(OriginalLedgerName, "'", "&quot;")

    ParentGroupName = Replace(ParentGroupName, "&", "&amp;")
    ParentGroupName = Replace(ParentGroupName, "'", "&apos;")
    ParentGroupName = Replace(ParentGroupName, "'", "&lt;")
    ParentGroupName = Replace(ParentGroupName, "'", "&gt;")
    ParentGroupName = Replace(ParentGroupName, "'", "&quot;")

    xmlString = modXml.AddOpeningTags(company, "All Masters") & xmlString
    xmlString = xmlString & modXml.ConvertLedgerToXML(TransactionType, LedgerName, OriginalLedgerName, ParentGroupName, OpeningBalance)
    xmlString = xmlString & modXml.AddClosingTags()
    If modGeneral.strMsg = "1" Then
        Dim FSO As New FileSystemObject
        Dim ts As TextStream
        Set ts = FSO.OpenTextFile(SetupFolder & "\LEDGER.xml", ForAppending, True)
        ts.Write xmlString
        ts.Close
    End If
    CreateAccount = ModTallyInterface.SendData(xmlString)
End Function




'******************************************************************************
'*****************************  OTHER FUNCTIONS   *****************************
'******************************************************************************

Private Function TrimSpaces(stringValue As String) As String

Dim status As Boolean
Dim temp1 As String
Dim temp2 As String

Dim I As Integer
I = 1
status = True

While status
    If Mid(stringValue, I, 1) = 32 Or Asc(Mid(stringValue, I, 1)) = 13 Or Asc(Mid(stringValue, I, 1)) = 10 Then
        temp1 = Trim(Mid(stringValue, 1, I))
        temp2 = Trim(Mid(stringValue, I))
        stringValue = temp1 & temp2
        I = 1
    End If
    If I = Len(stringValue) Then
        status = False
    End If
    I = I + 1
Wend
TrimSpaces = stringValue
End Function

Private Sub DisplayResult(Value As String)
    DoEvents
    'lblResult.Caption = vbCrLf & vbCrLf & "Processing ..." & vbCrLf & vbCrLf & Value
    lblResult.Caption = vbCrLf & vbCrLf & Value

End Sub


Public Function FillRecords(lvRecord As ListView)
    Dim COUNT As Integer
    Dim lst As ListItem
    lvRecord.ListItems.CLEAR
    For COUNT = 0 To UBound(TVchType) - 1
        Set lst = lvRecord.ListItems.Add(, , TVchType(COUNT))
        rec.Open "select count(*) as no from vch where txn_typ = '" & VchType(COUNT) & "'", CON, adOpenForwardOnly, adLockReadOnly
        lst.SubItems(1) = CStr(rec("no"))
        rec.Close
    Next
    For COUNT = 0 To UBound(TIVchType) - 1
        Set lst = lvRecord.ListItems.Add(, , TIVchType(COUNT))
        rec.Open "select count(*) as no from inl where txn_typ = '" & IVchType(COUNT) & "'", CON, adOpenForwardOnly, adLockReadOnly
        lst.SubItems(1) = CStr(rec("no"))
        rec.Close
    Next
End Function





'Inventory vouchers
Public Function CreateInventoryVoucherWithAcc(VoucherType As String, strParameterValue As String) As Long
'Tally Voucher Types
'PUR - Purchase
'SAL - Sales
'SRN - Credit Note
'PRT - Debit Note
'STI -
'STO -
    If VoucherType = "STI" Then
        CreateOutwards VoucherType, strParameterValue
        Exit Function
    ElseIf VoucherType = "OPS" Then
        Exit Function
    ElseIf VoucherType = "STO" Then
        Exit Function
    End If
    Dim TallyVchType As String
    Dim xmlstring1 As String
    Dim rec1 As New ADODB.Recordset
    Dim recInv As New ADODB.Recordset
    
    Dim cramt As Double
    Dim dramt As Double

    Dim sql1 As String
    Dim sql2 As String
    Dim RESULT As Long
    Dim PartyLedgerName As String
    Dim aid As String
    Dim IsDeemedPositive As String
    
    Dim AC1 As String, AC2 As String
    Dim PrdValAmt As Double, PrdDisAmt As Double, PrdTaxAmt As Double
    
    Dim CsdSpdAmt As Double, InlTaxAmt As Double, ExdAmt As Double, PkgCstAmt As Double
    Dim FrtAmt As Double, RndOffAmt As Double, NetAmt As Double
    
    Dim DisB As String, InvTax As String
    Dim DisA As String, InlTax As String, Exd As String, TxnRm1 As String
    Dim TxnRm2 As String, RndOff As String
    Dim T As Double
    Dim strTaxType As String
    Dim VATTaxLedgerName As String
    strTaxType = "Output"
    RESULT = 0

    Select Case VoucherType
    Case "PUR"
        strTaxType = "Input"
        TallyVchType = "Purchase"
    Case "SAL"
        strTaxType = "Output"
        TallyVchType = "Sales"
    Case "SRN"
        TallyVchType = "Credit Note"
    Case "PRT"
        TallyVchType = "Debit Note"

    End Select
    
    
    
           
    On Error GoTo ErrHandler
    
    sql1 = ""
    sql1 = "select INL.inv_idy, INL.INV_DAT, INL.txn_typ, sum(inv.col_amt) as prd_val, sum(INV.DIS_AMT) as prd_dis, SUM(INV.ADD_AMT) as add_dis, " & _
            " a2.act_idy as caid, a1.act_idy as daid,  " & _
            " sum(inv.tax_amt) as prd_tax,  INL.grs_amt, INL.spd_pcg, INL.spd_amt, INL.csd_pcg, INL.csd_amt, INL.exd_pcg, " & _
            " INL.exd_amt, INL.tax_pcg, INL.tax_amt as inl_tax_amt, INL.net_amt, INL.frt_amt, INL.pkg_cst, INL.rou_off , INL.cst_idy, INL.act_idy, a1.act_nme as AC1, a2.act_nme as AC2," & _
            " concat(concat(concat('Tax Output @ ' , trim(SUBSTR(INL.exd_pcg,5,2))), '%'),'A/c' ) as exd, " & _
            " 'VAT " & strTaxType & "' as invtax, concat(concat(concat('Tax Output @ ' , trim(SUBSTR(INL.tax_pcg,5,2))), '%'),'A/c' )  as inltax, CONCAT(trim(txn_rm1) , ' A/c')  as txnrm1," & _
            " CONCAT(trim(txn_rm2) , ' A/c') as txnrm2, 'Discount A/c (A)' as disA, 'Discount A/c (B)' as disB, " & _
            " 'Round Off A/c'  as roundoff  " & _
            " from inl, inv, act  a1, act  a2 " & _
            " where INL.inv_idy = inv.inv_idy and INL.cst_idy = a1.act_idy and INL.act_idy = a2.act_idy " & _
            " and INL.txn_typ='" & VoucherType & "' and INL.inv_idy in (" & strParameterValue & ") and inv.stk_typ ='Salable' "
    sql1 = sql1 & " group by INL.inv_idy, INL.INV_DAT, INL.txn_typ, INL.cst_idy, INL.act_idy, INL.grs_amt, INL.spd_pcg,  INL.spd_amt, INL.csd_pcg, INL.csd_amt,  INL.exd_pcg, INL.exd_amt, INL.tax_pcg, " & _
            " INL.tax_amt,  INL.net_amt, INL.frt_amt, INL.pkg_cst,INL.rou_off,a1.act_nme, a2.act_nme, txn_rm1, txn_rm2 , a1.act_idy, a2.act_idy"
        If rec.State = 1 Then rec.Close
        rec.Open sql1, CON, adOpenKeyset, adLockPessimistic
    Dim direct As String
    
    

    
    Dim I As Integer
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            xmlString = ""
            I = I + 1
        
        
        AC1 = Trim(rec("AC1"))
        AC2 = Trim(rec("AC2"))
        
        AC1 = Replace(AC1, "&", "&amp;")
        AC1 = Replace(AC1, "'", "&apos;")
        AC1 = Replace(AC1, "'", "&lt;")
        AC1 = Replace(AC1, "'", "&gt;")
        AC1 = Replace(AC1, "'", "&quot;")
        
        AC2 = Replace(AC2, "&", "&amp;")
        AC2 = Replace(AC2, "'", "&apos;")
        AC2 = Replace(AC2, "'", "&lt;")
        AC2 = Replace(AC2, "'", "&gt;")
        AC2 = Replace(AC2, "'", "&quot;")

        
        PrdValAmt = rec("prd_val")
        PrdDisAmt = rec("prd_dis") + rec("add_dis")
        CsdSpdAmt = rec("csd_amt") + rec("spd_amt")
        PrdValAmt = PrdValAmt - PrdDisAmt - CsdSpdAmt
        PrdTaxAmt = rec("prd_tax")
        
        DisB = rec("disB")
        InvTax = rec("invtax")
        DisA = rec("disA")
        InlTax = rec("inltax")
        Exd = rec("exd")
        TxnRm1 = rec("txnrm1")
        TxnRm2 = rec("txnrm2")
        RndOff = rec("roundoff")
        
        InlTaxAmt = rec("inl_tax_amt")
        ExdAmt = rec("exd_amt")
        PkgCstAmt = rec("pkg_cst")
        FrtAmt = rec("frt_amt")
        RndOffAmt = rec("rou_off")
        NetAmt = rec("net_amt")
        If NetAmt > 0 Then
'------------------------------------------------------------------------------
'------------------------------  ledger Creation  -----------------------------
'------------------------------------------------------------------------------
            RESULT = CreateAccount("Create", DisB, DisB, "Expenses(Indirect)", 0)
            If RESULT <> 0 Then GoTo ErrHandler
            If (PrdTaxAmt > 0) Then RESULT = CreateAccount("Create", InvTax, InvTax, "Duties & Taxes", 0)
            RESULT = CreateAccount("Create", DisA, DisA, "Expenses(Indirect)", 0)
            If InlTaxAmt > 0 Then RESULT = CreateAccount("Create", InlTax, InlTax, "Duties & Taxes", 0)
            If ExdAmt > 0 Then RESULT = CreateAccount("Create", Exd, Exd, "Duties & Taxes", 0)
            If PkgCstAmt > 0 Then RESULT = CreateAccount("Create", TxnRm1, TxnRm1, "Duties & Taxes", 0)
            If FrtAmt > 0 Then RESULT = CreateAccount("Create", TxnRm2, TxnRm2, "Duties & Taxes", 0)
            RESULT = CreateAccount("Create", RndOff, RndOff, "Expenses(Indirect)", 0)
'------------------------------------------------------------------------------
'------------------------------  Ledger Creation  -----------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------  Voucher Creation -----------------------------
'------------------------------------------------------------------------------
            PrdDisAmt = PrdDisAmt
            CsdSpdAmt = CsdSpdAmt
            FrtAmt = -FrtAmt
            PrdTaxAmt = PrdTaxAmt
            InlTaxAmt = InlTaxAmt
            PrdValAmt = PrdValAmt
            
            NetAmt = -NetAmt
            ExdAmt = -ExdAmt
            PkgCstAmt = PkgCstAmt
            RndOffAmt = RndOffAmt '(PrdValAmt + NetAmt + FrtAmt + PrdTaxAmt + InlTaxAmt + ExdAmt + PkgCstAmt) '
           
            
            If VoucherType = "PUR" Or VoucherType = "PRT" Then
                PrdDisAmt = -PrdDisAmt
                CsdSpdAmt = -CsdSpdAmt
                FrtAmt = -FrtAmt
                PrdTaxAmt = -PrdTaxAmt
                PrdValAmt = -PrdValAmt
                InlTaxAmt = -InlTaxAmt

                NetAmt = -NetAmt
                ExdAmt = -ExdAmt
                PkgCstAmt = -PkgCstAmt
                RndOffAmt = -RndOffAmt
    
                PartyLedgerName = Trim(AC1)
                aid = Trim(rec("caid"))
                If FrtAmt <> 0 Then direct = "Direct Incomes"
            End If
            If VoucherType = "SAL" Or VoucherType = "SRN" Then
                PartyLedgerName = Trim(AC1)
                aid = Trim(rec("daid"))
                If FrtAmt <> 0 Then direct = "Direct Expenses"
        
            End If
            DisplayResult "         Voucher Type - " & TallyVchType & "  ID - " & Trim(rec("inv_idy"))
            xmlString = ""
            xmlString = xmlString & modXml.AddOpeningTags(company, "All Masters", TallyVchType)
            xmlString = xmlString & modXml.VoucherOpeningTag("Create", TallyVchType, Trim(rec("inv_idy")) & "-" & aid, Trim(rec("inv_idy")), Trim(PartyLedgerName), Trim(rec("inv_dat")), "Yes", Trim(rec("inv_idy")))
            xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(AC1), Format(NetAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
'            If PrdTaxAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(InvTax), Format(PrdTaxAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            If InlTaxAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(InlTax), Format(InlTaxAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            If ExdAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(Exd), Format(ExdAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            If PkgCstAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(TxnRm1), Format(PkgCstAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            If FrtAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(direct), Format(FrtAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            If RndOffAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(RndOff), Format(RndOffAmt, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
            If CsdSpdAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(DisA), -CsdSpdAmt, IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "Yes", "No"))
           If recInv.State = 1 Then recInv.Close
            recInv.Open "SELECT INV.PRD_IDY, PRD.PRD_NME, INV.PRD_RAT , INV.PRD_QTY, INV.TRD_TAX, (INV.PRD_RAT * INV.PRD_QTY) AS AMOUNT, INV.ADD_AMT , INV.DIS_AMT , TAX_AMT, INV.STK_TYP  FROM INV, PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND PRD_QTY>0 AND INV.PRD_RAT <> 0 AND TRIM(INV.INV_IDY)='" & Trim(rec("inv_idy")) & "' ", CON, adOpenKeyset, adLockPessimistic
'            INV.PRD_RAT, INV.PRD_QTY,INV.PRD_RAT* INV.PRD_QTY, INV.DIS_AMT, INV.ADD_AMT,
            Dim TrdAmt As Double
            Dim PrdRate As Double
            Dim trdtax As Double
            
            While Not recInv.EOF
            If IsNull(recInv!TRD_TAX) <> True And recInv!TRD_TAX > 0 Then
                If VoucherType = "SAL" Or VoucherType = "SRN" Then
                    VATTaxLedgerName = "Output VAT @ " & STR(recInv!TRD_TAX) & "%"
                Else
                    VATTaxLedgerName = "Input VAT @ " & STR(recInv!TRD_TAX) & "%"
                End If
            Else
                If VoucherType = "SAL" Or VoucherType = "SRN" Then
                    VATTaxLedgerName = "VAT Output"
                Else
                    VATTaxLedgerName = "VAT Input"
                End If
            End If
                TrdAmt = recInv!amount
                PrdDisAmt = IIf(IsNull(recInv!DIS_AMT) = True, 0, recInv!DIS_AMT) + IIf(IsNull(recInv!ADD_AMT) = True, 0, recInv!ADD_AMT)
                trdtax = recInv!TAX_AMT
'                TrdAmt = TrdAmt - IIf(IsNull(recInv!DIS_AMT) = True, 0, recInv!DIS_AMT)
'                TrdAmt = TrdAmt - IIf(IsNull(recInv!ADD_AMT) = True, 0, recInv!ADD_AMT)
'                TrdAmt = TrdAmt + IIf(IsNull(recInv!TAX_AMT) = True, 0, recInv!TAX_AMT)
                PrdRate = TrdAmt / recInv!prd_qty
                If TrdAmt > 0 Then
                TrdAmt = IIf(Trim(recInv!stk_typ) = "Salable", TrdAmt, -(TrdAmt))
                TrdAmt = IIf((VoucherType = "PUR" Or VoucherType = "PRT"), -(TrdAmt), TrdAmt)
                End If
                IsDeemedPositive = IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "Yes", "No")
'                IsDeemedPositive = IIf(Trim(recInv!stk_typ) = "Salable", "No", "Yes")
                xmlString = xmlString & ConvertInventoryToXML(Trim(recInv!PRD_NME), TrdAmt, recInv!prd_qty, PrdRate, recInv!TRD_TAX, IIf((VoucherType = "PUR" Or VoucherType = "PRT"), recInv!TAX_AMT, -recInv!TAX_AMT), VATTaxLedgerName, Trim(AC2), IsDeemedPositive)
'                xmlString = xmlString & ConvertInventoryToXML(Trim(AC2), recInv!amount - IIf(IsNull(recInv!ADD_AMT) = True, 0, recInv!ADD_AMT) - IIf(IsNull(recInv!DIS_AMT) = True, 0, recInv!DIS_AMT), recInv!PRD_QTY, recInv!PRD_RAT, recInv!TRD_TAX, IIf((VoucherType = "PUR" Or VoucherType = "PRT"), recInv!TAX_AMT, -recInv!TAX_AMT), VATTaxLedgerName, IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "No", "Yes"))
'                xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(AC2), Format(rec!amount, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "Yes", "No"))
                PrdDisAmt = IIf((VoucherType = "PUR" Or VoucherType = "PRT"), -(PrdDisAmt), PrdDisAmt)
             If PrdDisAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(DisB), -PrdDisAmt, IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "Yes", "No"))
                PrdDisAmt = IIf((VoucherType = "PUR" Or VoucherType = "PRT"), -(PrdDisAmt), PrdDisAmt)
                trdtax = IIf((VoucherType = "PUR" Or VoucherType = "PRT"), -(trdtax), trdtax)
                
               If trdtax <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(VATTaxLedgerName, Format(trdtax, "############.00"), IIf((VoucherType = "PUR" Or VoucherType = "PRT"), "Yes", "No"))
                recInv.MoveNext
            Wend
'            xmlString = xmlString & modXml.InventoryClosingTag()
            xmlString = xmlString & modXml.VoucherClosingTag()
            xmlString = xmlString & modXml.AddClosingTags()
        
            If modGeneral.strMsg = "1" Then
                Dim FSO As New FileSystemObject
                Dim ts As TextStream
                Set ts = FSO.OpenTextFile(SetupFolder & "\" & Trim(rec("inv_idy")) & ".xml", ForAppending, True)
                ts.Write xmlString
                ts.Close
            End If
            RESULT = ModTallyInterface.SendData(xmlString)
        End If
        rec.MoveNext
    Loop
End If
    '------------------------------------
        If rec.State = 1 Then rec.Close
        CreateInventoryVoucherWithAcc = RESULT
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If RESULT <> 0 Then RESULT = err.Number
    CreateInventoryVoucherWithAcc = RESULT
End Function


Public Function CreateOutwards(VoucherType As String, strParameterValue As String)


    Dim TallyVchType As String
    Dim xmlstring1 As String
    Dim rec1 As New ADODB.Recordset
    Dim recInv As New ADODB.Recordset

    Dim cramt As Double
    Dim dramt As Double

    Dim sql1 As String
    Dim sql2 As String
    Dim RESULT As Long
    Dim PartyLedgerName As String
    Dim aid As String
    Dim IsDeemedPositive As String

    Dim AC1 As String, AC2 As String
    Dim PrdValAmt As Double, PrdDisAmt As Double, PrdTaxAmt As Double

    Dim CsdSpdAmt As Double, InlTaxAmt As Double, ExdAmt As Double, PkgCstAmt As Double
    Dim FrtAmt As Double, RndOffAmt As Double, NetAmt As Double

    Dim strTaxType As String
    strTaxType = "Output"
    RESULT = 0

    Select Case VoucherType
    Case "STI"
        strTaxType = "Input"
        TallyVchType = "Stock Journal"
    End Select

    sql1 = ""
    sql1 = "select INL.inv_idy, INL.INV_DAT, INL.txn_typ " & _
            " from inl" & _
            " where INL.txn_typ='" & VoucherType & "' and INL.inv_idy in (" & strParameterValue & ")  "
        If rec.State = 1 Then rec.Close
        rec.Open sql1, CON, adOpenKeyset, adLockPessimistic
        PartyLedgerName = ""
    Do While Not rec.EOF
'------------------------------------------------------------------------------
'------------------------------  Voucher Creation -----------------------------
'------------------------------------------------------------------------------
            DisplayResult "         Voucher Type - " & TallyVchType & "  ID - " & Trim(rec("inv_idy"))
            xmlString = ""
            xmlString = xmlString & modXml.AddOpeningTags(company, "All Masters", TallyVchType)
            xmlString = xmlString & modXml.VoucherOpeningTag("Create", TallyVchType, Trim(rec("inv_idy")), Trim(rec("inv_idy")), Trim(PartyLedgerName), Trim(rec("inv_dat")), "Yes", Trim(rec("inv_idy")))
            If recInv.State = 1 Then recInv.Close
            recInv.Open "SELECT INV.PRD_IDY, PRD.PRD_NME, INV.PRD_RAT , INV.PRD_QTY, INV.TRD_TAX, (INV.PRD_RAT * INV.PRD_QTY) AS AMOUNT, INV.ADD_AMT , INV.DIS_AMT , TAX_AMT, INV.STK_TYP  FROM INV, PRD WHERE PRD.PRD_IDY = INV.PRD_IDY AND PRD_QTY>0 AND INV.PRD_RAT <> 0 AND TRIM(INV.INV_IDY)='" & Trim(rec("inv_idy")) & "' ", CON, adOpenKeyset, adLockPessimistic
            Dim TrdAmt As Double
            Dim PrdRate As Double
            Dim trdtax As Double

            While Not recInv.EOF
                xmlString = xmlString & ConvertOutwardToXML(Trim(recInv!PRD_NME), recInv!amount, recInv!prd_qty, recInv!PRD_rat)
                recInv.MoveNext
            Wend
            xmlString = xmlString & modXml.VoucherClosingTag()
            xmlString = xmlString & modXml.AddClosingTags()

            If modGeneral.strMsg = "1" Then
                Dim FSO As New FileSystemObject
                Dim ts As TextStream
                Set ts = FSO.OpenTextFile(SetupFolder & "\" & Trim(rec("inv_idy")) & ".xml", ForAppending, True)
                ts.Write xmlString
                ts.Close
            End If
            RESULT = ModTallyInterface.SendData(xmlString)
        rec.MoveNext
    Loop

    '------------------------------------
        If rec.State = 1 Then rec.Close
        CreateOutwards = RESULT
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If RESULT <> 0 Then RESULT = err.Number
    CreateOutwards = RESULT
End Function
