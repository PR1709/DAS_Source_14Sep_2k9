Attribute VB_Name = "modXml"
Option Explicit
Dim Company As String


Public Function AddOpeningTags(Company As String, ReportName As String, Optional AccountType As String) As String
    Dim xmlString As String
    
    xmlString = "<ENVELOPE>"
    xmlString = xmlString & "<HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>"
    xmlString = xmlString & "<BODY>"
    xmlString = xmlString & "<IMPORTDATA>"
    xmlString = xmlString & "<REQUESTDESC>"
    xmlString = xmlString & "<REPORTNAME>" & ReportName & "</REPORTNAME>"
    xmlString = xmlString & "<STATICVARIABLES>"
    xmlString = xmlString & "<SVCURRENTCOMPANY>" & Company & "</SVCURRENTCOMPANY>"
'    xmlstring = xmlstring & "<ACCOUNTTYPE>" & AccountType & "</ACCOUNTTYPE>"
    xmlString = xmlString & "</STATICVARIABLES>"
    xmlString = xmlString & "</REQUESTDESC>"
    xmlString = xmlString & "<REQUESTDATA>"
    xmlString = xmlString & "<TALLYMESSAGE xmlns:UDF='TallyUDF'>"

    AddOpeningTags = xmlString
End Function

Public Function AddClosingTags() As String
    Dim xmlString As String
    xmlString = xmlString & "</TALLYMESSAGE>"
    xmlString = xmlString & "</REQUESTDATA>"
    xmlString = xmlString & "</IMPORTDATA>"
    xmlString = xmlString & "</BODY>"
    xmlString = xmlString & "</ENVELOPE>"
    
    AddClosingTags = xmlString
End Function

Public Function CompanyList() As String
Dim xmlString As String
    xmlString = xmlString & "<ENVELOPE>"
    xmlString = xmlString & "<HEADER>"
    xmlString = xmlString & "<TALLYREQUEST>Export Data</TALLYREQUEST>"
    xmlString = xmlString & "</HEADER>"
    xmlString = xmlString & "<BODY>"
    xmlString = xmlString & "<EXPORTDATA>"
    xmlString = xmlString & "<REQUESTDESC>"
    xmlString = xmlString & "<REPORTNAME>List of Companies</REPORTNAME>"
    xmlString = xmlString & "</REQUESTDESC>"
    xmlString = xmlString & "</EXPORTDATA>"
    xmlString = xmlString & "</BODY>"
    xmlString = xmlString & "</ENVELOPE>"
    CompanyList = xmlString
End Function

Public Function ConvertGroupToXML(TransactionType As String, GroupName As String, OriginalGroupName As String, ParentGroupName As String, Optional IsDeemedPositive As String, Optional isRevenue As String) As String
    Dim xmlString As String
'    xmlString = AddOpeningTags(Company)
    xmlString = xmlString & "<GROUP NAME='" & GroupName & "'  ACTION='" & TransactionType & "'>"
    xmlString = xmlString & "<NAME.LIST>"
    xmlString = xmlString & "<NAME>" & OriginalGroupName & "</NAME>"
    xmlString = xmlString & "</NAME.LIST>"
    xmlString = xmlString & "<PARENT>" & ParentGroupName & "</PARENT>"
    If IsDeemedPositive = "" Then IsDeemedPositive = "No"
    If isRevenue = "" Then isRevenue = "No"
    xmlString = xmlString & "<ISDEEMEDPOSITIVE>" & IsDeemedPositive & "</ISDEEMEDPOSITIVE>"
    xmlString = xmlString & "<ISREVENUE>" & isRevenue & "</ISREVENUE>"
    xmlString = xmlString & "<ISSUBLEDGER>No</ISSUBLEDGER>"
    xmlString = xmlString & "<TRACKNEGATIVEBALANCES>No</TRACKNEGATIVEBALANCES>"
    xmlString = xmlString & "<ISBILLWISEON>No</ISBILLWISEON>"
    xmlString = xmlString & "<ISCOSTCENTRESON>No</ISCOSTCENTRESON>"
    xmlString = xmlString & "<AFFECTSSTOCK>No</AFFECTSSTOCK>"
    
'    xmlString = xmlString & "<ISADDABLE>No</ISADDABLE>"
'    xmlString = xmlString & "<AFFECTSGROSSPROFIT>No</AFFECTSGROSSPROFIT>"
'    xmlString = xmlString & "<ISCONDENSED>No</ISCONDENSED>"
'    xmlString = xmlString & "<SORTPOSITION>60</SORTPOSITION>"
    xmlString = xmlString & "</GROUP>"
'    xmlString = xmlString & AddClosingTags()

    ConvertGroupToXML = xmlString
End Function

Public Function ConvertLedgerToXML(TransactionType As String, LedgerName As String, OriginalLedgerName As String, ParentGroupName As String, OpeningBalance As Double) As String
    Dim xmlString As String

    xmlString = xmlString & "<LEDGER NAME='" & LedgerName & "'  ACTION='" & TransactionType & "'>"
    xmlString = xmlString & "<NAME.LIST>"
    xmlString = xmlString & "<NAME>" & OriginalLedgerName & "</NAME>"
    xmlString = xmlString & "</NAME.LIST>"
    xmlString = xmlString & "<PARENT>" & ParentGroupName & "</PARENT>"
    xmlString = xmlString & "<OPENINGBALANCE>" & OpeningBalance & "</OPENINGBALANCE>"
    xmlString = xmlString & "<CURRENCYNAME>Rs.</CURRENCYNAME>"
'    xmlString = xmlString & "<ISBILLWISEON>No</ISBILLWISEON>"
'    xmlString = xmlString & "<AFFECTSSTOCK>No</AFFECTSSTOCK>"
'    xmlString = xmlString & "<USEFORVAT>No</USEFORVAT>"
'    xmlString = xmlString & "<TAXCLASSIFICATIONNAME />"
'    xmlString = xmlString & "<TAXTYPE />"
'    xmlString = xmlString & "<RATEOFTAXCALCULATION />"
'    xmlString = xmlString & "<GSTTYPE />"
'    xmlString = xmlString & "<SERVICECATEGORY />"
'    xmlString = xmlString & "<EXCISEDUTYTYPE />"
'    xmlString = xmlString & "<TRADERLEDNATUREOFPURCHASE />"
'    xmlString = xmlString & "<TDSDEDUCTEETYPE />"
'    xmlString = xmlString & "<TDSRATENAME />"
'    xmlString = xmlString & "<LEDGERFBTCATEGORY />"
'    xmlString = xmlString & "<ISCOSTCENTRESON>No</ISCOSTCENTRESON>"
'    xmlString = xmlString & "<ISINTERESTON>No</ISINTERESTON>"
'    xmlString = xmlString & "<ALLOWINMOBILE>No</ALLOWINMOBILE>"
'    xmlString = xmlString & "<ISCONDENSED>No</ISCONDENSED>"
'    xmlString = xmlString & "<FORPAYROLL>No</FORPAYROLL>"
'    xmlString = xmlString & "<INTERESTONBILLWISE>No</INTERESTONBILLWISE>"
'    xmlString = xmlString & "<OVERRIDEINTEREST>No</OVERRIDEINTEREST>"
'    xmlString = xmlString & "<OVERRIDEADVINTEREST>No</OVERRIDEADVINTEREST>"
'    xmlString = xmlString & "<IGNORETDSEXEMPT>No</IGNORETDSEXEMPT>"
'    xmlString = xmlString & "<ISTCSAPPLICABLE>No</ISTCSAPPLICABLE>"
'    xmlString = xmlString & "<ISTDSAPPLICABLE>No</ISTDSAPPLICABLE>"
'    xmlString = xmlString & "<ISFBTAPPLICABLE>No</ISFBTAPPLICABLE>"
'    xmlString = xmlString & "<ISGSTAPPLICABLE>No</ISGSTAPPLICABLE>"
'    xmlString = xmlString & "<SHOWINPAYSLIP>No</SHOWINPAYSLIP>"
'    xmlString = xmlString & "<USEFORGRATUITY>No</USEFORGRATUITY>"
'    xmlString = xmlString & "<FORSERVICETAX>No</FORSERVICETAX>"
'    xmlString = xmlString & "<ISINPUTCREDIT>No</ISINPUTCREDIT>"
'    xmlString = xmlString & "<ISEXEMPTED>No</ISEXEMPTED>"
'    xmlString = xmlString & "<ISABATEMENTAPPLICABLE>No</ISABATEMENTAPPLICABLE>"
'    xmlString = xmlString & "<TDSDEDUCTEEISSPECIALRATE>No</TDSDEDUCTEEISSPECIALRATE>"
'    xmlString = xmlString & "<AUDITED>No</AUDITED>"
    'xmlString = xmlString & "<SORTPOSITION>1000</SORTPOSITION>"
    xmlString = xmlString & "</LEDGER>"


    ConvertLedgerToXML = xmlString
End Function

Public Function ConvertUnitToXML(TransactionType As String, UnitName As String) As String
    Dim xmlString As String

    xmlString = xmlString & "<UNIT NAME='" & UnitName & "'  ACTION='" & TransactionType & "'>"
    xmlString = xmlString & "<NAME>" & UnitName & "</NAME>"
    xmlString = xmlString & "<ISSIMPLEUNIT>Yes</ISSIMPLEUNIT>"
    xmlString = xmlString & "<FORPAYROLL>No</FORPAYROLL>"
    xmlString = xmlString & "</UNIT>"


    ConvertUnitToXML = xmlString
End Function

Public Function ConvertStockGroupToXML(TransactionType As String, StockGroupName As String, OriginalStockGroupName As String) As String
    Dim xmlString As String

    xmlString = xmlString & "<STOCKGROUP NAME='" & StockGroupName & "'  ACTION='" & TransactionType & "'>"
    xmlString = xmlString & "<NAME.LIST>"
    xmlString = xmlString & "<NAME>" & OriginalStockGroupName & "</NAME>"
    xmlString = xmlString & "</NAME.LIST>"
    xmlString = xmlString & "<PARENT />"
    xmlString = xmlString & "<ISADDABLE>Yes</ISADDABLE>"
    xmlString = xmlString & "</STOCKGROUP>"

    ConvertStockGroupToXML = xmlString
End Function

Public Function ConvertStockItemToXML(TransactionType As String, SKUName As String, SKUDescription As String, ParentGroupName As String, UOM As String, _
OpeningBalance As Double, ReorderLevel As Double, MinimumLevel As Double, PurchaseRate As Double, SalesRate As Double) As String
    Dim xmlString As String

    xmlString = xmlString & "<STOCKITEM NAME='" & Trim(SKUName) & "'  ACTION='" & TransactionType & "'>"
    xmlString = xmlString & "<NAME.LIST>"
    xmlString = xmlString & "<NAME>" & Trim(SKUName) & "</NAME>"
    xmlString = xmlString & "</NAME.LIST>"
'    xmlString = xmlString & "<ADDITIONALNAME.LIST>"
'    xmlString = xmlString & "<ADDITIONALNAME>" & Trim(SKUDescription) & "</ADDITIONALNAME>"
'    xmlString = xmlString & "</ADDITIONALNAME.LIST>"
    xmlString = xmlString & "<PARENT>" & Trim(ParentGroupName) & "</PARENT>"
    xmlString = xmlString & "<BASEUNITS>" & Trim(UOM) & "</BASEUNITS>"
    xmlString = xmlString & "<OPENINGBALANCE>" & OpeningBalance & "</OPENINGBALANCE>"
    xmlString = xmlString & "<RATE>" & SalesRate & "</RATE>"
    xmlString = xmlString & "<RATE>" & PurchaseRate & "</RATE>"
    xmlString = xmlString & "<REORDERBASE>" & ReorderLevel & "</REORDERBASE>"
    'xmlString = xmlString & "<OPENINGVALUE></OPENINGVALUE>"
    'xmlString = xmlString & "<OPENINGRATE></OPENINGRATE>"
    xmlString = xmlString & "<MINIMUMORDERBASE>MinimumLevel</MINIMUMORDERBASE>"
''''''    xmlString = xmlString & "<BATCHALLOCATIONS.LIST>"
''''''    xmlString = xmlString & "<NAME></NAME>"
''''''    xmlString = xmlString & "<BATCHNAME></BATCHNAME>"
''''''    xmlString = xmlString & "<GODOWNNAME></GODOWNNAME>"
''''''    xmlString = xmlString & "<MFDON></MFDON>"
''''''    xmlString = xmlString & "<OPENINGBALANCE></OPENINGBALANCE>"
''''''    xmlString = xmlString & "<OPENINGVALUE></OPENINGVALUE>"
''''''    xmlString = xmlString & "<OPENINGRATE></OPENINGRATE>"
''''''    xmlString = xmlString & "</BATCHALLOCATIONS.LIST>"
'    xmlString = xmlString & "<STANDARDPRICELIST.LIST>"
'    xmlString = xmlString & "<DATE></DATE>"
'    xmlString = xmlString & "</STANDARDPRICELIST.LIST>"
'    xmlString = xmlString & "<STANDARDCOSTLIST.LIST>"
'    xmlString = xmlString & "<DATE></DATE>"
'    xmlString = xmlString & "</STANDARDCOSTLIST.LIST>"
    xmlString = xmlString & "</STOCKITEM>"


    ConvertStockItemToXML = xmlString
End Function


Public Function VoucherOpeningTag(TransactionType As String, VoucherType As String, RemoteId As String, VoucherId As String, PartyLedgerName As String, VoucherDate As String) As String
    Dim xmlString As String

Dim strDate As String
strDate = Year(VoucherDate) & Format(Month(VoucherDate), "00") & Format(Day(VoucherDate), "00")

    xmlString = xmlString & "<VOUCHER REMOTEID='" & RemoteId & "' VCHTYPE='" & VoucherType & "' ACTION='" & TransactionType & "'>"
    xmlString = xmlString & "<ISOPTIONAL>No</ISOPTIONAL>"
    xmlString = xmlString & "<USEFORGAINLOSS>No</USEFORGAINLOSS>"
    xmlString = xmlString & "<USEFORCOMPOUND>No</USEFORCOMPOUND>"
    xmlString = xmlString & "<VOUCHERTYPENAME>" & VoucherType & "</VOUCHERTYPENAME>"
    xmlString = xmlString & "<DATE>" & strDate & "</DATE>"
    xmlString = xmlString & "<EFFECTIVEDATE>" & strDate & "</EFFECTIVEDATE>"
    xmlString = xmlString & "<ISCANCELLED>No</ISCANCELLED>"
    xmlString = xmlString & "<USETRACKINGNUMBER>Yes</USETRACKINGNUMBER>"
    xmlString = xmlString & "<ISPOSTDATED>No</ISPOSTDATED>"
    xmlString = xmlString & "<ISINVOICE>No</ISINVOICE>"
    xmlString = xmlString & "<DIFFACTUALQTY>No</DIFFACTUALQTY>"
    xmlString = xmlString & "<ASPAYSLIP>No</ASPAYSLIP>"
    xmlString = xmlString & "<GUID>" & RemoteId & "</GUID>"
    xmlString = xmlString & "<REFERENCE>" & VoucherId & "</REFERENCE>"
    xmlString = xmlString & "<PARTYLEDGERNAME>" & PartyLedgerName & "</PARTYLEDGERNAME>"
    xmlString = xmlString & "<NARRATION />"
    VoucherOpeningTag = xmlString
End Function

Public Function VoucherClosingTag()
    Dim xmlString As String
    xmlString = xmlString & "<INVOICEDELNOTES.LIST />"
    xmlString = xmlString & "</VOUCHER>"
    VoucherClosingTag = xmlString
End Function

Public Function ConvertVoucherToXML(LedgerName As String, amount As String, IsDeemedPositive As String, Optional RemoveZeroEntries As String)
    Dim xmlString As String
    xmlString = xmlString & "<LEDGERENTRIES.LIST>"
    If RemoveZeroEntries = "" Then RemoveZeroEntries = "No"
    xmlString = xmlString & "<REMOVEZEROENTRIES>" & RemoveZeroEntries & "</REMOVEZEROENTRIES>"
    xmlString = xmlString & "<ISDEEMEDPOSITIVE>" & IsDeemedPositive & "</ISDEEMEDPOSITIVE>"
    xmlString = xmlString & "<LEDGERFROMITEM>No</LEDGERFROMITEM>"
    xmlString = xmlString & "<LEDGERNAME>" & LedgerName & "</LEDGERNAME>"
    xmlString = xmlString & "<AMOUNT>" & amount & "</AMOUNT>"
    xmlString = xmlString & "</LEDGERENTRIES.LIST>"

    ConvertVoucherToXML = xmlString
End Function



