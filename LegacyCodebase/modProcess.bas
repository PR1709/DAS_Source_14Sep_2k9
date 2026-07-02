Attribute VB_Name = "modProcess"
Option Explicit
'Public con As New Connection
Public Company As String
Public lblResult As Label
Public lblProcess As Label

Dim xmlString As String
Dim rec As New ADODB.Recordset

Public IVchType(7)  As String
Public VchType(5)   As String

Public TIVchType(7)  As String
Public TVchType(5)   As String

Public Function AddZephyrCompany(cmbZephyrCompany As ComboBox, strFile As String)
    Dim fso As New FileSystemObject
    Dim txtStream As TextStream
    Dim str As String
    
    Set txtStream = fso.OpenTextFile(strFile, ForReading)
    txtStream.ReadLine
    cmbZephyrCompany.CLEAR
    While Not txtStream.AtEndOfLine
        str = txtStream.ReadLine()
        str = Mid(str, 2, InStr(1, str, "@") - 2)
        cmbZephyrCompany.AddItem str
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
    If strResponse <> "" And nodeList.Length <> 0 Then
        For nCount = 0 To nodeList.Length - 1
            cmbTallyCompany.AddItem nodeList.Item(nCount).Text
        Next
    Else
'-----------if comes in text format
        Dim bool As Boolean
        bool = True
        Dim s As String
        Dim i As Integer
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
    Dim result As Long
    result = 0
    On Error GoTo ErrHandler
    DoEvents
    
    Dim fso As FileSystemObject
    Dim txtStream As TextStream
    Dim strLine As String
    Dim LedgerName As String
    Dim IsDeemedPostivie As String
    Dim isRevenue As String
    Dim temploc As Integer
    Dim temp As String
    
    
    Set fso = New FileSystemObject
    Set txtStream = fso.OpenTextFile(App.Path & "\PrimaryGroups.txt", ForReading)
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
        xmlString = modXml.AddOpeningTags(Company, "All Masters") & xmlString
        xmlString = xmlString & modXml.ConvertGroupToXML("Create", LedgerName, LedgerName, "", IsDeemedPostivie, isRevenue)
        xmlString = xmlString & modXml.AddClosingTags()
        result = ModTallyInterface.SendData(xmlString)
    Loop
    CreatePrimaryGroups = result
    Exit Function
ErrHandler:
    CreatePrimaryGroups = result
End Function

Public Function CreateGroups() As Long
'Error Handler
    On Error Resume Next
    Dim result As Long
    rec.Open "select grp_idy, grp_nme, grp_clf from grp order by grp_idy", con, adOpenForwardOnly, adLockReadOnly
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
                xmlString = modXml.AddOpeningTags(Company, "All Masters") & xmlString
                xmlString = xmlString & modXml.ConvertGroupToXML("Create", UCase(Trim(rec("grp_nme"))), UCase(Trim(rec("grp_nme"))), Trim(rec("grp_clf")))
                'xmlstring = xmlstring & modXml.ConvertGroupToXML("Create", Replace(rec("grp_nme"), "&", "-"), Replace(rec("grp_nme"), "&", "-"), Replace(TrimSpaces(rec("grp_clf")), "&", "-"))
                xmlString = xmlString & modXml.AddClosingTags()
                result = ModTallyInterface.SendData(xmlString)
            End If
            rec.MoveNext
        Loop
    End If
    If rec.State = 1 Then rec.Close
    CreateGroups = result
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If result <> 0 Then result = err.Number
    CreateGroups = result
End Function

Public Function CreateAccounts() As Long
'Error Handler
    On Error GoTo ErrHandler
    Dim result As Long
    Dim rec1 As New ADODB.Recordset
    Dim amt As Double
    
    result = 0
    rec.Open "select act_idy, act_nme, grp_nme from act,grp where grp.grp_idy = act.grp_idy order by act_idy", con, adOpenForwardOnly, adLockReadOnly
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            amt = 0
            'OPN - Opening Balance
            rec1.Open "select sum(txn_amt) as opn from vch where crt_idy = '" & rec("act_idy") & "' and txn_typ='OPN'", con, adOpenStatic, adLockOptimistic
            If Not rec1.EOF Then amt = rec1("opn")
            rec1.Close
            If amt = 0 Then
                rec1.Open "select sum(txn_amt) as opn from vch where dbt_idy = '" & rec("act_idy") & "' and txn_typ='OPN'", con, adOpenStatic, adLockOptimistic
                If Not rec1.EOF Then amt = rec1("opn")
                rec1.Close
            End If
            'OPS - Opening Stock Balance
            If amt = 0 Then
                rec1.Open "select sum(net_amt) as opn from inl where act_idy = '" & rec("act_idy") & "' and txn_typ='OPS'", con, adOpenForwardOnly, adLockReadOnly
                If Not rec1.EOF Then amt = rec1("opn")
                rec1.Close
            End If
            If amt = 0 Then
                rec1.Open "select sum(net_amt) as opn from inl where cst_idy = '" & rec("act_idy") & "' and txn_typ='OPS'", con, adOpenForwardOnly, adLockReadOnly
                If Not rec1.EOF Then amt = rec1("opn")
                rec1.Close
            End If
            If rec1.State = 1 Then rec1.Close
            result = CreateAccount("Create", Replace(rec("act_nme"), "&", "-"), Replace(rec("act_nme"), "&", "-"), Replace(rec("grp_nme"), "&", "-"), amt)
            rec.MoveNext
        Loop
    End If
    result = CreateAccount("Create", "VAT Output", "VAT Output", "Duties & Taxes", 0)
    If rec.State = 1 Then rec.Close
    CreateAccounts = result
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If result <> 0 Then result = err.Number
    CreateAccounts = result
End Function

Public Function CreateUOM() As Long
'Error Handler
    On Error GoTo ErrHandler
    Dim result As Long
    result = 0
    rec.Open "select distinct prd_uom from prd order by prd_uom", con, adOpenForwardOnly, adLockReadOnly
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            DisplayResult "         UOM - " & rec("prd_uom")
            xmlString = ""
            xmlString = modXml.AddOpeningTags(Company, "All Masters") & xmlString
            xmlString = xmlString & modXml.ConvertUnitToXML("Create", Replace(rec("prd_uom"), "&", "-"))
            xmlString = xmlString & modXml.AddClosingTags()
            result = ModTallyInterface.SendData(xmlString)
            rec.MoveNext
        Loop
    End If
    If rec.State = 1 Then rec.Close
    CreateUOM = result
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If result <> 0 Then result = err.Number
    CreateUOM = result
End Function

Public Function CreateStockGroup() As Long
'Error Handler
    On Error GoTo ErrHandler
    Dim result As Long
    result = 0
    rec.Open "select distinct grp_cls from pgr order by grp_cls", con, adOpenForwardOnly, adLockReadOnly
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            DisplayResult "         Stock Group - " & rec("grp_cls")
            xmlString = ""
            xmlString = modXml.AddOpeningTags(Company, "All Masters") & xmlString
            xmlString = xmlString & modXml.ConvertStockGroupToXML("Create", Replace(rec("grp_cls"), "&", "-"), Replace(rec("grp_cls"), "&", "-"))
            xmlString = xmlString & modXml.AddClosingTags()
            result = ModTallyInterface.SendData(xmlString)
            rec.MoveNext
        Loop
    End If
    If rec.State = 1 Then rec.Close
    CreateStockGroup = result
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If result <> 0 Then result = err.Number
    CreateStockGroup = result
End Function

Public Function CreateStockItem() As Long
'Error Handler
    On Error GoTo ErrHandler
    Dim result As Long
    result = 0
    rec.Open "select distinct  prd_nme, pgr.grp_cls, prd_uom, pur_rat, sal_rat, mrp_rat, prd.cur_lev, prd.reo_lev, prd.saf_lev from pgr, prd where pgr.grp_idy = prd.grp_idy order by prd_nme", con, adOpenForwardOnly, adLockReadOnly
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            DisplayResult "         Stock Item (SKU) - " & rec("prd_nme")
            xmlString = ""
            xmlString = modXml.AddOpeningTags(Company, "All Masters") & xmlString
            xmlString = xmlString & modXml.ConvertStockItemToXML("Create", Replace(rec("prd_nme"), "&", "-"), Replace(rec("prd_nme"), "&", "-"), Replace(rec("grp_cls"), "&", "-"), Replace(rec("prd_uom"), "&", "-"), CDbl(rec("cur_lev")), CDbl(rec("reo_lev")), CDbl(rec("saf_lev")), CDbl(rec("pur_rat")), CDbl(rec("sal_rat")))
            xmlString = xmlString & modXml.AddClosingTags()
            result = ModTallyInterface.SendData(xmlString)
            rec.MoveNext
        Loop
    End If
    If rec.State = 1 Then rec.Close
    CreateStockItem = result
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If result <> 0 Then result = err.Number
    CreateStockItem = result
End Function

Public Function CreateAccountingVoucher(VoucherType As String) As Long
'Tally Voucher Types
'JRN - Journal
'PMT - Payment
'RCT - Receipt
'Error Handler


    Dim TallyVchType As String
    Dim xmlstring1 As String
    Dim sql1 As String
    Dim result As Long
    Dim DrAmount As String
    Dim CrAmount As String
    Dim DrAccount As String
    Dim CrAccount As String
    Dim PartyLedgerName As String
    Dim IsDeemedPositive As String
    
    result = 0
    If VoucherType = "OPN" Then Exit Function
    Select Case VoucherType
    Case "RCT"
        TallyVchType = "Receipt"
    Case "DBN"
        TallyVchType = "Debit Note"
    Case "JRN"
        TallyVchType = "Journal"
    Case "PMT"
        TallyVchType = "Payment"
    Case "CRN"
        TallyVchType = "Credit Note"
    End Select

    On Error GoTo ErrHandler
    sql1 = ""
    sql1 = "select txn_idy, txn_dat, txn_typ, txn_amt, a1.act_nme as dr_nme, a2.act_nme as cr_nme  " & _
            " from vch, act as a1, act as a2 " & _
            " where vch.dbt_idy = a1.act_idy And vch.crt_idy = a2.act_idy  " & _
            " and vch.txn_typ='" & VoucherType & "' order by txn_idy"
    rec.Open sql1, con, adOpenForwardOnly, adLockReadOnly
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            xmlString = ""
            DrAmount = "-0.00"
            CrAmount = "0.00"
            DrAmount = CStr(-1 * CDbl(rec("txn_amt")))
            CrAmount = CStr(CDbl(rec("txn_amt")))
            CrAccount = Trim(rec("cr_nme"))
            DrAccount = Trim(rec("Dr_nme"))
            If VoucherType = "RCT" Or VoucherType = "DBN" Then
               PartyLedgerName = DrAccount
            End If
            If VoucherType = "PMT" Or VoucherType = "CRN" Then
                PartyLedgerName = CrAccount
            End If
            If VoucherType = "JRN" Then
                PartyLedgerName = ""
            End If
            DisplayResult "         Voucher Type - " & TallyVchType & "  ID - " & Trim(rec("txn_idy"))
            xmlString = xmlString & modXml.AddOpeningTags(Company, "All Masters", TallyVchType)
            xmlString = xmlString & modXml.VoucherOpeningTag("Create", TallyVchType, "6593251k-803u-731g-953p-597825398422-" & Trim(rec("txn_idy")), Trim(rec("txn_idy")), PartyLedgerName, Trim(rec("txn_dat")))
            xmlString = xmlString & modXml.ConvertVoucherToXML(DrAccount, DrAmount, "Yes")
            xmlString = xmlString & modXml.ConvertVoucherToXML(CrAccount, CrAmount, "No")
            xmlString = xmlString & modXml.VoucherClosingTag()
            xmlString = xmlString & modXml.AddClosingTags()
            result = ModTallyInterface.SendData(xmlString)
            rec.MoveNext
        Loop
    End If
    '------------------------------------
    If rec.State = 1 Then rec.Close
    CreateAccountingVoucher = result
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If result <> 0 Then result = err.Number
    CreateAccountingVoucher = result
End Function


'Inventory vouchers
Public Function CreateInventoryVoucher(VoucherType As String, FromDate As Date, ToDate As Date) As Long
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
    Dim result As Long
    Dim PartyLedgerName As String
    
    
    Dim AC1 As String, AC2 As String
    Dim PrdValAmt As Double, PrdDisAmt As Double, PrdTaxAmt As Double
    
    Dim CsdSpdAmt As Double, InlTaxAmt As Double, ExdAmt As Double, PkgCstAmt As Double
    Dim FrtAmt As Double, RndOffAmt As Double, NetAmt As Double
    
    Dim DisB As String, InvTax As String
    Dim DisA As String, InlTax As String, Exd As String, TxnRm1 As String
    Dim TxnRm2 As String, RndOff As String
    Dim T As Double
    
    result = 0

    Select Case VoucherType
    Case "PUR"
        TallyVchType = "Purchase"
    Case "SAL"
        TallyVchType = "Sales"
    Case "SRN"
        TallyVchType = "Credit Note"
    Case "PRT"
        TallyVchType = "Debit Note"
'    Case "STI"
'        TallyVchType = ""
'    Case "STO"
'        TallyVchType = ""
    End Select
    
    
    
           
    On Error GoTo ErrHandler
    sql1 = ""
    sql1 = "select inl.inv_idy, inl.inv_dat, inl.txn_typ, sum(inv.col_amt) as prd_val, sum(inv.dis_amt) as prd_dis," & _
            " sum(inv.tax_amt) as prd_tax,  inl.grs_amt, inl.spd_pcg, inl.spd_amt, inl.csd_pcg, inl.csd_amt, inl.exd_pcg," & _
            " inl.exd_amt, inl.tax_pcg, inl.tax_amt as inl_tax_amt, inl.net_amt, inl.frt_amt, inl.pkg_cst, inl.rou_off ," & _
            " inl.cst_idy, inl.act_idy, a1.act_nme as AC1, a2.act_nme as AC2, 'Tax Output @ ' + trim(str(inl.exd_pcg,5,2)) + '%' as exd," & _
            " 'VAT Output' as invtax, 'Tax Output @ ' + trim(str(inl.tax_pcg,5,2)) + '%' as inltax, " & _
            " trim(txn_rm1) + ' A/c'  as txnrm1, trim(txn_rm2) + ' A/c' as txnrm2, 'Discount A/c (A)' as disA, 'Discount A/c (B)' as disB, " & _
            " 'Round Off A/c'  as roundoff  " & _
            " from inl, inv, act as a1, act as a2 " & _
            " where inl.inv_idy = inv.inv_idy and inl.cst_idy = a1.act_idy and inl.act_idy = a2.act_idy " & _
            " and inl.txn_typ='" & VoucherType & "' " & _
            " and inl.inv_dat  => date( " & Year(FromDate) & "," & Month(FromDate) & "," & Day(FromDate) & ") " & _
            " and inl.inv_dat  => date( " & Year(ToDate) & "," & Month(ToDate) & "," & Day(ToDate) & ") " & _
            " group by inl.inv_idy, inl.inv_dat, inl.txn_typ, inl.cst_idy, inl.act_idy, inl.grs_amt, inl.spd_pcg, " & _
            " inl.spd_amt, inl.csd_pcg, inl.csd_amt,  inl.exd_pcg, inl.exd_amt, inl.tax_pcg, inl.tax_amt, " & _
            " inl.net_amt, inl.frt_amt, inl.pkg_cst,inl.rou_off"
        rec.Open sql1, con, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rec.EOF Then
        Do While Not rec.EOF
            DoEvents
            xmlString = ""
            
        
        
        AC1 = rec("AC1")
        AC2 = rec("AC2")
        PrdValAmt = rec("prd_val")
        PrdDisAmt = rec("prd_dis")
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
'------------------------------------------------------------------------------
'------------------------------  ledger Creation  -----------------------------
'------------------------------------------------------------------------------
        result = CreateAccount("Create", DisB, DisB, "Expenses(Indirect)", 0)
        If result <> 0 Then GoTo ErrHandler
        If (PrdTaxAmt > 0) Then result = CreateAccount("Create", InvTax, InvTax, "Duties & Taxes", 0)
        result = CreateAccount("Create", DisA, DisA, "Expenses(Indirect)", 0)
        If InlTaxAmt > 0 Then result = CreateAccount("Create", InlTax, InlTax, "Duties & Taxes", 0)
        If ExdAmt > 0 Then result = CreateAccount("Create", Exd, Exd, "Duties & Taxes", 0)
        If PkgCstAmt > 0 Then result = CreateAccount("Create", TxnRm1, TxnRm1, "Duties & Taxes", 0)
        If FrtAmt > 0 Then result = CreateAccount("Create", TxnRm2, TxnRm2, "Duties & Taxes", 0)
        result = CreateAccount("Create", RndOff, RndOff, "Expenses(Indirect)", 0)
'------------------------------------------------------------------------------
'------------------------------  Ledger Creation  -----------------------------
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
'------------------------------  Voucher Creation -----------------------------
'------------------------------------------------------------------------------
        
        If VoucherType = "PUR" Or VoucherType = "PRT" Then
            PrdDisAmt = -PrdDisAmt
            CsdSpdAmt = -CsdSpdAmt
            FrtAmt = -FrtAmt
            NetAmt = -NetAmt
            
            PrdValAmt = PrdValAmt
            PrdTaxAmt = PrdTaxAmt
            InlTaxAmt = InlTaxAmt
            ExdAmt = ExdAmt
            PkgCstAmt = PkgCstAmt

            PartyLedgerName = AC1
            RndOffAmt = FrtAmt + NetAmt + PrdValAmt + PrdTaxAmt + InlTaxAmt + ExdAmt + PkgCstAmt
'            RndOffAmt = PrdDisAmt + CsdSpdAmt + FrtAmt + NetAmt + PrdValAmt + PrdTaxAmt + InlTaxAmt + ExdAmt + PkgCstAmt
'            If RndOffAmt > 0 Then
                RndOffAmt = RndOffAmt
'            Else
'                RndOffAmt = RndOffAmt
'            End If
            PartyLedgerName = AC2

        End If
        If VoucherType = "SAL" Or VoucherType = "SRN" Then
            PrdDisAmt = PrdDisAmt
            CsdSpdAmt = CsdSpdAmt
            FrtAmt = FrtAmt
            NetAmt = NetAmt
            
            PrdValAmt = -PrdValAmt
            PrdTaxAmt = -PrdTaxAmt
            InlTaxAmt = -InlTaxAmt
            ExdAmt = -ExdAmt
            PkgCstAmt = -PkgCstAmt
            RndOffAmt = FrtAmt + NetAmt + PrdValAmt + PrdTaxAmt + InlTaxAmt + ExdAmt + PkgCstAmt
'            RndOffAmt = PrdDisAmt + CsdSpdAmt + FrtAmt + NetAmt + PrdValAmt + PrdTaxAmt + InlTaxAmt + ExdAmt + PkgCstAmt
'            If RndOffAmt > 0 Then
                RndOffAmt = -RndOffAmt
'            Else
'                RndOffAmt = -RndOffAmt
'            End If
            PartyLedgerName = AC2
            
        End If
            DisplayResult "         Voucher Type - " & TallyVchType & "  ID - " & Trim(rec("inv_idy"))
            xmlString = ""
            xmlString = xmlString & modXml.AddOpeningTags(Company, "All Masters", TallyVchType)
            xmlString = xmlString & modXml.VoucherOpeningTag("Create", TallyVchType, "6593251k-803u-731g-953p-597825398422-" & Trim(rec("inv_idy")), Trim(rec("inv_idy")), Trim(PartyLedgerName), Trim(rec("inv_dat")))
            
            xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(AC1), Format(PrdValAmt, "############.00"), IIf(PrdValAmt > 0, "Yes", "No"))
'            If PrdDisAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(DisB), Format(PrdDisAmt, "############.00"), IIf(PrdDisAmt <= 0, "Yes", "No"))
            If PrdTaxAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(InvTax), Format(PrdTaxAmt, "############.00"), IIf(PrdTaxAmt <= 0, "Yes", "No"))
'            If CsdSpdAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(DisA), Format(CsdSpdAmt, "############.00"), IIf(CsdSpdAmt <= 0, "Yes", "No"))
            If InlTaxAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(InlTax), Format(InlTaxAmt, "############.00"), IIf(InlTaxAmt <= 0, "Yes", "No"))
            If ExdAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(Exd), Format(ExdAmt, "############.00"), IIf(ExdAmt <= 0, "Yes", "No"))
            If PkgCstAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(TxnRm1), Format(PkgCstAmt, "############.00"), IIf(PkgCstAmt <= 0, "Yes", "No"))
            If FrtAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(TxnRm2), Format(FrtAmt, "############.00"), IIf(FrtAmt <= 0, "Yes", "No"))
            If RndOffAmt <> 0 Then xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(RndOff), Format(RndOffAmt, "############.00"), IIf(RndOffAmt <= 0, "Yes", "No"))
            xmlString = xmlString & modXml.ConvertVoucherToXML(Trim(AC2), Format(NetAmt, "############.00"), IIf(NetAmt > 0, "Yes", "No"))
            
            xmlString = xmlString & modXml.VoucherClosingTag()
            xmlString = xmlString & modXml.AddClosingTags()
            result = ModTallyInterface.SendData(xmlString)
            rec.MoveNext
        Loop
    End If
    '------------------------------------
    If rec.State = 1 Then rec.Close
    CreateInventoryVoucher = result
    Exit Function
ErrHandler:
    If rec.State = 1 Then rec.Close
    If result <> 0 Then result = err.Number
    CreateInventoryVoucher = result
End Function


Private Function CreateAccount(TransactionType As String, LedgerName As String, OriginalLedgerName As String, ParentGroupName As String, OpeningBalance As Double) As Long
    Dim xmlString As String
    DoEvents
    DisplayResult "         Account - " & LedgerName
    xmlString = ""
    xmlString = modXml.AddOpeningTags(Company, "All Masters") & xmlString
    xmlString = xmlString & modXml.ConvertLedgerToXML(TransactionType, LedgerName, OriginalLedgerName, ParentGroupName, OpeningBalance)
    xmlString = xmlString & modXml.AddClosingTags()
    CreateAccount = ModTallyInterface.SendData(xmlString)
End Function




'******************************************************************************
'*****************************  OTHER FUNCTIONS   *****************************
'******************************************************************************

Private Function TrimSpaces(stringValue As String) As String

Dim status As Boolean
Dim temp1 As String
Dim temp2 As String

Dim i As Integer
i = 1
status = True

While status
    If Mid(stringValue, i, 1) = 32 Or Asc(Mid(stringValue, i, 1)) = 13 Or Asc(Mid(stringValue, i, 1)) = 10 Then
        temp1 = Trim(Mid(stringValue, 1, i))
        temp2 = Trim(Mid(stringValue, i))
        stringValue = temp1 & temp2
        i = 1
    End If
    If i = Len(stringValue) Then
        status = False
    End If
    i = i + 1
Wend
TrimSpaces = stringValue
End Function

Private Sub DisplayResult(Value As String)
    DoEvents
    'lblResult.Caption = vbCrLf & vbCrLf & "Processing ..." & vbCrLf & vbCrLf & Value
    lblResult.Caption = vbCrLf & vbCrLf & Value

End Sub


Public Function FillRecords(lvRecord As ListView)
    Dim count As Integer
    Dim lst As ListItem
    
    lvRecord.ListItems.CLEAR
    For count = 0 To UBound(TVchType) - 1
        Set lst = lvRecord.ListItems.Add(, , TVchType(count))
        rec.Open "select count(*) as no from vch where txn_typ = '" & VchType(count) & "'", con, adOpenForwardOnly, adLockReadOnly
        lst.SubItems(1) = CStr(rec("no"))
        rec.Close
    Next
    For count = 0 To UBound(TIVchType) - 1
        Set lst = lvRecord.ListItems.Add(, , TIVchType(count))
        rec.Open "select count(*) as no from inl where txn_typ = '" & IVchType(count) & "'", con, adOpenForwardOnly, adLockReadOnly
        lst.SubItems(1) = CStr(rec("no"))
        rec.Close
    Next


End Function



