VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{09E60688-07A2-11CF-BD27-10005AE6AFFA}#1.0#0"; "RRW32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm M_MDI 
   BackColor       =   &H8000000F&
   Caption         =   "DAS Version 4.01.01 Rel. - Distributor Automation System, Ver.Rel.04-Sep-2009"
   ClientHeight    =   8640
   ClientLeft      =   165
   ClientTop       =   -810
   ClientWidth     =   12735
   Icon            =   "M_MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar STS_BAR 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8265
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16351
            MinWidth        =   3175
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1693
            MinWidth        =   1499
            TextSave        =   "09/04/2009"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1589
            MinWidth        =   1589
            TextSave        =   "6:04 PM"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   10875
      Left            =   0
      ScaleHeight     =   10815
      ScaleWidth      =   12675
      TabIndex        =   1
      Top             =   0
      Width           =   12735
      Begin VB.Image Image1 
         Height          =   1425
         Left            =   10440
         Picture         =   "M_MDI.frx":08CA
         Stretch         =   -1  'True
         Top             =   6120
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consoft Systems Pvt Ltd"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   10005
         TabIndex        =   2
         Top             =   7560
         Width           =   2235
      End
   End
   Begin RrocxLib.RRReport RRReport1 
      Left            =   1920
      Top             =   1560
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      WindowControlBox=   0   'False
      WindowBorderStyle=   0
      WindowTitle     =   ""
      PrintFileName   =   ""
      ReportName      =   ""
      ReportLibrary   =   ""
      MemoFileName    =   ""
      Printer         =   ""
      Port            =   ""
      MasterTable     =   ""
      Filter          =   ""
      RelatedTables   =   ""
      WindowParentHandle=   0
      DataDirectory   =   ""
      ReportDirectory =   ""
      ImageDirectory  =   ""
      LastErrorCode   =   0
      LastErrorString =   ""
      LastErrorPage   =   0
      SuppressTitle   =   0   'False
      Parameters      =   ""
      DBContainer     =   ""
      SortFields      =   ""
      GroupFields     =   ""
      ReportSortFields=   ""
      ReportGroupFields=   ""
      ControlFile     =   ""
      SuitableGroupFields=   ""
      SuitableSortFields=   ""
      ReportsRTString =   ""
      RelatedTablesString=   ""
      ParametersString=   ""
      ReportParametersString=   ""
      StatusFileName  =   ""
      MasterIndex     =   ""
      LowScope        =   ""
      HighScope       =   ""
      IndexExtension  =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   1560
   End
   Begin VB.Menu M_FIL 
      Caption         =   "&File"
      Enabled         =   0   'False
      Tag             =   "File"
      Begin VB.Menu M_FIL_CAT 
         Caption         =   "&Channels"
         Enabled         =   0   'False
         Tag             =   "Categories"
      End
      Begin VB.Menu M_FIL_LOC 
         Caption         =   "&Locations"
      End
      Begin VB.Menu M_FIL_SMN 
         Caption         =   "&Salesmen"
      End
      Begin VB.Menu M_FIL_MAR 
         Caption         =   "&Markets"
      End
      Begin VB.Menu M_FIL_HYP1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu M_FIL_GRP 
         Caption         =   "&Group Accounts"
      End
      Begin VB.Menu M_FIL_ACT 
         Caption         =   "&Accounts && Parties"
         Shortcut        =   {F2}
      End
      Begin VB.Menu M_FIL_HYP2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu M_FIL_PGR 
         Caption         =   "Product Cat&egories"
      End
      Begin VB.Menu M_FIL_PRD 
         Caption         =   "&Products (SKU)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu M_FIL_HYP4 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu M_FIL_BDT 
         Caption         =   "Beats, Cal-PJP, Coms, HVOs, Key A/cs, Schemes, Interfaces "
         Begin VB.Menu M_FIL_SCH 
            Caption         =   "&1 Schemes SKU,Discounts and CashOff"
            Index           =   0
            Shortcut        =   ^H
         End
      End
      Begin VB.Menu M_FIL_HYP6 
         Caption         =   "-"
      End
      Begin VB.Menu M_FIL_DAS_CRS 
         Caption         =   "DAS C&OLA Replenishment Xtract"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu M_FIL_HYP5 
         Caption         =   "Day Clos&Ure n Data XChange"
         Index           =   0
         Shortcut        =   ^G
      End
      Begin VB.Menu M_FIL_HYP7 
         Caption         =   "-"
      End
      Begin VB.Menu M_FIL_EXT 
         Caption         =   "e&Xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu M_COR 
      Caption         =   "&Core"
      Enabled         =   0   'False
      Tag             =   "Core"
      Begin VB.Menu M_COR_PUR_SSL 
         Caption         =   "&Purchases"
         Enabled         =   0   'False
         Shortcut        =   {F8}
         Tag             =   "Purchases"
      End
      Begin VB.Menu M_COR_SAL 
         Caption         =   "&Sales"
         Enabled         =   0   'False
         Tag             =   "Sales"
         Begin VB.Menu M_COR_FRE_SAL 
            Caption         =   "&Sales - Invoices"
            Index           =   0
            Shortcut        =   {F4}
            Tag             =   "Credit Sales"
         End
         Begin VB.Menu M_COR_SAL_RSB 
            Caption         =   "&Van Stock Issue"
            Index           =   0
            Shortcut        =   {F5}
            Tag             =   "Ready Stock Bills"
         End
         Begin VB.Menu M_COR_VAN_VVR 
            Caption         =   "Va&n Variance Check"
            Index           =   0
         End
         Begin VB.Menu mnu_bill 
            Caption         =   "&Collection Entry"
            Shortcut        =   {F6}
            Tag             =   "Invoices vs Receipts"
         End
         Begin VB.Menu M_COR_SAL_LDS 
            Caption         =   "&Load Sheet"
            Index           =   0
            Shortcut        =   {F7}
            Tag             =   "Load Sheet"
         End
         Begin VB.Menu M_COR_PSH_SAL 
            Caption         =   "&Push Cart Sales"
         End
      End
      Begin VB.Menu M_COR_STK 
         Caption         =   "Stoc&ks"
         Enabled         =   0   'False
         Tag             =   "Stocks"
         Begin VB.Menu M_COR_STK_OPS 
            Caption         =   "&a Opening Stock"
            Tag             =   "Opening Stock"
         End
         Begin VB.Menu M_COR_STK_CRN 
            Caption         =   "&b Credit Notes - Sale Return"
            Shortcut        =   ^R
            Tag             =   "Credit Notes"
         End
         Begin VB.Menu M_COR_STK_DRN 
            Caption         =   "&c Debit Notes - Purchase Return"
            Shortcut        =   ^D
            Tag             =   "Debit Notes"
         End
         Begin VB.Menu M_COR_STK_HYP1 
            Caption         =   "-"
         End
         Begin VB.Menu M_COR_STK_STI 
            Caption         =   "&d Unplanned Receipt"
            Shortcut        =   ^I
            Tag             =   "Transfer Inward"
         End
         Begin VB.Menu M_COR_STK_STO 
            Caption         =   "&e Unplanned Issue"
            Index           =   0
            Shortcut        =   ^O
            Tag             =   "Transfer Outward"
         End
         Begin VB.Menu M_COR_STK_ADJ_JRN 
            Caption         =   "&f Stock Adj.Journal"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu M_COR_STK_COV 
            Caption         =   "&g Stock Conversion"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu M_COR_HYP7 
         Caption         =   "-"
      End
      Begin VB.Menu M_MNU_SHL_SNP1 
         Caption         =   "&Returnables Accounting - 2Way Packs"
         Begin VB.Menu M_COR_SHL_TRS 
            Caption         =   "&a Returnables - Accounting"
            Shortcut        =   ^A
         End
         Begin VB.Menu M_MNU_SHL_INW 
            Caption         =   "&b Returnables - Transactions"
            Index           =   0
            Shortcut        =   ^E
         End
         Begin VB.Menu M_MNU_SHL_SNP 
            Caption         =   "&c Returnables - Shell Identification"
            Shortcut        =   ^W
         End
         Begin VB.Menu M_COR_TWO_REP 
            Caption         =   "&d Returnables - Reports"
            Begin VB.Menu M_COR_TWO_REP_ESH 
               Caption         =   "&a Returnables (RGB) Stock on Hand"
               Index           =   0
            End
         End
      End
      Begin VB.Menu M_COR_HYP5 
         Caption         =   "Te&lecom"
         Index           =   0
         Visible         =   0   'False
         Begin VB.Menu TCO_ENT 
            Caption         =   "&Activation Targets"
            Index           =   0
         End
      End
      Begin VB.Menu SHL_SEP 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MNU_CMP 
         Caption         =   "&Orders, Load Sheets, Verifications, Xrfs"
         Enabled         =   0   'False
         Index           =   0
         Tag             =   "Supplier Order"
         Begin VB.Menu OLXRF 
            Caption         =   "&1 Sales Orders Entry"
            Index           =   0
            Shortcut        =   ^S
         End
         Begin VB.Menu OLXRF 
            Caption         =   "&3 Sales Order Load Sheet"
            Index           =   3
            Shortcut        =   ^{F7}
         End
      End
      Begin VB.Menu M_COR_UPS 
         Caption         =   "&Update Stocks"
         Enabled         =   0   'False
         Shortcut        =   ^U
         Tag             =   "Update Stocks"
      End
   End
   Begin VB.Menu M_ACT 
      Caption         =   "&Accounts"
      Enabled         =   0   'False
      Tag             =   "Accounts"
      Begin VB.Menu M_ACT_RCT 
         Caption         =   "&Receipts"
         Enabled         =   0   'False
         Shortcut        =   ^{F2}
         Tag             =   "Receipts"
      End
      Begin VB.Menu M_ACT_HYP1 
         Caption         =   "-"
      End
      Begin VB.Menu M_ACT_PMT 
         Caption         =   "&Payments"
         Enabled         =   0   'False
         Shortcut        =   ^{F3}
         Tag             =   "Payments"
      End
      Begin VB.Menu M_ACT_HYP2 
         Caption         =   "-"
      End
      Begin VB.Menu M_ACT_JRN 
         Caption         =   "&Journals"
         Enabled         =   0   'False
         Tag             =   "Journals"
         Begin VB.Menu M_ACT_VCH_JVN 
            Caption         =   "&A/c Journals"
            Enabled         =   0   'False
            Shortcut        =   ^{F6}
            Tag             =   "Journal Entries"
         End
         Begin VB.Menu M_ACT_VCH_CRN 
            Caption         =   "&Credit Notes"
            Enabled         =   0   'False
            Shortcut        =   ^{F4}
            Tag             =   "Credit Note"
         End
         Begin VB.Menu M_ACT_VCH_DRN 
            Caption         =   "&Debit Notes"
            Enabled         =   0   'False
            Shortcut        =   ^{F5}
            Tag             =   "Debit Note"
         End
         Begin VB.Menu M_ACT_VCH_OPB 
            Caption         =   "&Opening Balance"
            Enabled         =   0   'False
            Shortcut        =   ^{F8}
            Tag             =   "Opening Balance"
         End
      End
   End
   Begin VB.Menu M_AST 
      Caption         =   "Ass&ets"
      Enabled         =   0   'False
      Tag             =   "Addons"
      Begin VB.Menu M_FIL_ATT 
         Caption         =   "&Asset Type"
      End
      Begin VB.Menu M_FIL_ASM 
         Caption         =   "Asset &Detail"
      End
      Begin VB.Menu M_AST_INW 
         Caption         =   "&Inward Entry"
      End
      Begin VB.Menu M_AST_OUT 
         Caption         =   "&Outward Entry"
      End
   End
   Begin VB.Menu m_mir 
      Caption         =   "Anal&ysis"
      Enabled         =   0   'False
      Tag             =   "Analysis"
      Begin VB.Menu ANA_RET 
         Caption         =   "&Retailers"
         Enabled         =   0   'False
         Tag             =   "Retailers"
         Begin VB.Menu M_MIR_TOT 
            Caption         =   "&a Ranking"
         End
         Begin VB.Menu M_ANA_RET_LBD 
            Caption         =   "&b Last Billing Date"
         End
         Begin VB.Menu M_ANA_RET_RRN 
            Caption         =   "&c Retailers Route NB"
         End
         Begin VB.Menu M_ANA_RET_RSN 
            Caption         =   "&d Retailers Salesman NB"
         End
         Begin VB.Menu M_MIR_NPD 
            Caption         =   "&e Non Operational"
         End
         Begin VB.Menu M_MIR_RCG 
            Caption         =   "&f Categories Count"
         End
         Begin VB.Menu M_MIR_SOT 
            Caption         =   "&g Sales Pareto"
         End
         Begin VB.Menu M_MIR_CFR 
            Caption         =   "&h Coverage Frequency"
         End
         Begin VB.Menu M_MIR_YSR 
            Caption         =   "&i Yearly Sales"
         End
         Begin VB.Menu M_ANA_RET_SKU 
            Caption         =   "&j Retailers - SKU"
         End
      End
      Begin VB.Menu ANA_RET_HYP_1 
         Caption         =   "-"
      End
      Begin VB.Menu ANA_SKU 
         Caption         =   "&Products (SKU)"
         Enabled         =   0   'False
         Tag             =   "Products(SKU)"
         Begin VB.Menu M_MIR_PPD 
            Caption         =   "&a SKU Pareto"
         End
         Begin VB.Menu M_MIR_PPD_SEG 
            Caption         =   "&b SKU Segment Pareto"
         End
         Begin VB.Menu M_MIR_PPD_CAT 
            Caption         =   "&c SKU Category Pareto"
         End
         Begin VB.Menu M_MIR_PPV 
            Caption         =   "&d Sale Value Pareto"
            Index           =   0
         End
         Begin VB.Menu M_MIR_SPD 
            Caption         =   "&e Retailers Pareto"
         End
         Begin VB.Menu m_mir_cat 
            Caption         =   "&f Categories Pareto"
         End
         Begin VB.Menu m_mar_loc 
            Caption         =   "&g Locations Pareto"
         End
         Begin VB.Menu M_MIR_MAR 
            Caption         =   "&h Markets Pareto"
         End
         Begin VB.Menu M_MIR_ROU 
            Caption         =   "&i Routes Pareto"
            Index           =   0
         End
         Begin VB.Menu M_MIR_RAR 
            Caption         =   "&j Penetration Count"
         End
         Begin VB.Menu mnu_sup 
            Caption         =   "&k Sales Report - Splr"
         End
         Begin VB.Menu mnu_stk 
            Caption         =   "&l Sales Report - Stock"
         End
         Begin VB.Menu M_RPT_COR_STK_ANL_MVR 
            Caption         =   "&m Product Sales"
            Index           =   0
         End
         Begin VB.Menu M_MIR_YQP 
            Caption         =   "&n Yearly Sales - Units"
         End
         Begin VB.Menu M_MIR_YQC 
            Caption         =   "&o Yearly Sales - Cases"
         End
         Begin VB.Menu mnu_sup_b 
            Caption         =   "&p Actual Sales Report"
         End
      End
      Begin VB.Menu M_ANA_RPT_SEP3 
         Caption         =   "-"
      End
      Begin VB.Menu M_ANA_SPC_RPT 
         Caption         =   "&Specific Reports"
         Begin VB.Menu VAN_VAR 
            Caption         =   "&a Van Variance"
         End
         Begin VB.Menu SHL_VAR 
            Caption         =   "&b Shell Variance"
         End
         Begin VB.Menu VAN_SHL_VAR 
            Caption         =   "&c Van Shell Variance"
         End
         Begin VB.Menu SMN_CUM_BAL 
            Caption         =   "&d Smn Cumulative Balance"
         End
         Begin VB.Menu DLY_PRD_SUM 
            Caption         =   "&e Daily Productivity Summary"
         End
         Begin VB.Menu DIS_SCH_CPC 
            Caption         =   "&f Discount Scheme CPC"
         End
         Begin VB.Menu VEH_UTL 
            Caption         =   "&g Vehicle Utilization"
         End
         Begin VB.Menu VAN_CHL_STS 
            Caption         =   "&h Van Challans Status"
         End
         Begin VB.Menu DIS_STK_REP 
            Caption         =   "&i Distributors Stock Report"
            Index           =   0
         End
      End
      Begin VB.Menu M_ANA_RPT_SEP1 
         Caption         =   "-"
      End
      Begin VB.Menu M_ANA_PRG_RPT 
         Caption         =   "Progress &Reports"
         Begin VB.Menu M_ANA_PRG_CAT 
            Caption         =   "&a Category"
         End
         Begin VB.Menu M_ANA_PRG_LOC 
            Caption         =   "&b Location"
         End
         Begin VB.Menu M_ANA_PRG_MAR 
            Caption         =   "&c Market"
         End
         Begin VB.Menu M_ANA_PRG_SMN 
            Caption         =   "&d Salesman"
         End
         Begin VB.Menu M_ANA_PRG_ROU 
            Caption         =   "&e Route"
         End
         Begin VB.Menu M_ANA_PRG_RET 
            Caption         =   "&f Retailer"
         End
      End
      Begin VB.Menu M_ANA_RPT_SEP2 
         Caption         =   "-"
      End
      Begin VB.Menu M_ANA_COM_RPT 
         Caption         =   "&Combination Reports"
      End
   End
   Begin VB.Menu M_RPT 
      Caption         =   "&Reports"
      Enabled         =   0   'False
      Tag             =   "Reports"
      Begin VB.Menu M_RPT_COR 
         Caption         =   "&Core"
         Enabled         =   0   'False
         Tag             =   "Core(R)"
         Begin VB.Menu M_RPT_COR_SAL 
            Caption         =   "&Sales"
            Enabled         =   0   'False
            Tag             =   "Sales(R)"
            Begin VB.Menu M_RPT_COR_SAL_RPT 
               Caption         =   "&a Report - Date"
               Index           =   0
            End
         End
         Begin VB.Menu M_RPT_COR_HYP1 
            Caption         =   "-"
         End
         Begin VB.Menu M_RPT_COR_PUR 
            Caption         =   "&Purchases"
            Enabled         =   0   'False
            Tag             =   "Purchases(R)"
            Begin VB.Menu M_RPT_COR_PUR_SUM 
               Caption         =   "&a Report"
            End
            Begin VB.Menu M_RPT_COR_PUR_PRD 
               Caption         =   "&b Product"
            End
            Begin VB.Menu M_RPT_COR_PUR_RET 
               Caption         =   "&c Returns"
            End
            Begin VB.Menu M_RPT_OTH_A9S 
               Caption         =   "&d A-9 Statement"
            End
            Begin VB.Menu M_RPT_COR_PUR_PRD_TAX 
               Caption         =   "&e Tax Product"
            End
            Begin VB.Menu M_RPT_COR_PUR_PRD_TAX_TTL 
               Caption         =   "&f Total Tax Product"
            End
            Begin VB.Menu M_RPT_COR_PUR_MAR_RPT 
               Caption         =   "&g Margin Report"
               Index           =   0
            End
            Begin VB.Menu M_RPT_COR_PUR_TID_RPT 
               Caption         =   "&i Transit Brk-Lkg-Dmgs"
            End
         End
         Begin VB.Menu M_RPT_COR_HYP2 
            Caption         =   "-"
         End
         Begin VB.Menu M_RPT_COR_STK 
            Caption         =   "S&tocks"
            Enabled         =   0   'False
            Tag             =   "Stocks(R)"
            Begin VB.Menu M_RPT_COR_STK_ANL_AGD 
               Caption         =   "&a SKU - Aging - BBD-Exp Date"
               Index           =   0
            End
            Begin VB.Menu M_RPT_COR_STK_RGL_STL 
               Caption         =   "&m Stock Ledger - Detailed - by SKU"
               Index           =   0
            End
         End
         Begin VB.Menu ANA_RET_HYP_2 
            Caption         =   "-"
         End
         Begin VB.Menu ANA_OTH 
            Caption         =   "K&O Reports"
            Enabled         =   0   'False
            Tag             =   "Others"
            Begin VB.Menu M_MIR_MBR 
               Caption         =   "&a Retailer Sales and SGA Productivity"
               Index           =   0
            End
         End
      End
      Begin VB.Menu M_RPT_HYP2 
         Caption         =   "-"
      End
      Begin VB.Menu M_RPT_ACT 
         Caption         =   "&Accounts"
         Enabled         =   0   'False
         Tag             =   "Accounts(R)"
         Begin VB.Menu M_RPT_ACT_BBK 
            Caption         =   "&a Bank n Cash Book"
         End
         Begin VB.Menu M_RPT_ACT_SMM 
            Caption         =   "&b Day Summary"
         End
         Begin VB.Menu M_RPT_ACT_DCR 
            Caption         =   "&c Collection Report"
         End
         Begin VB.Menu M_RPT_ACT_CIB 
            Caption         =   "&d Collection Cheques n Breakup - 2"
         End
         Begin VB.Menu M_RPT_ACT_CRL 
            Caption         =   "&e Creditors' List"
         End
         Begin VB.Menu M_RPT_ACT_DLS 
            Caption         =   "&f Debtors' List"
         End
         Begin VB.Menu REC_STM 
            Caption         =   "&g Debtors' A/c"
         End
         Begin VB.Menu DebtorsAgeing 
            Caption         =   "&h Debtors' Ageing"
         End
         Begin VB.Menu dbt_agd 
            Caption         =   "&i Invoices Aging"
         End
         Begin VB.Menu M_RPT_ACT_DCR_SCR 
            Caption         =   "&j Day A/c OC Report"
         End
         Begin VB.Menu M_RPT_ACT_HYP3 
            Caption         =   "-"
         End
         Begin VB.Menu M_RPT_ACT_GLD 
            Caption         =   "&k General Ledger"
         End
         Begin VB.Menu M_RPT_ACT_ACR 
            Caption         =   "&l Registers"
            Begin VB.Menu M_RPT_ACT_RVL 
               Caption         =   "&a Receipts "
            End
            Begin VB.Menu M_RPT_ACT_PVL 
               Caption         =   "&b Payments"
            End
            Begin VB.Menu M_RPT_ACT_REG_HYP1 
               Caption         =   "-"
            End
            Begin VB.Menu M_RPT_ACT_CNR 
               Caption         =   "&c Credit Notes"
            End
            Begin VB.Menu M_RPT_ACT_DNR 
               Caption         =   "&d Debit Notes"
            End
            Begin VB.Menu M_RPT_ACT_REG_HYP2 
               Caption         =   "-"
            End
            Begin VB.Menu M_RPT_ACT_JVL 
               Caption         =   "&e Journals"
            End
            Begin VB.Menu M_RPT_ACT_OAL 
               Caption         =   "&f Opening Balance"
            End
         End
         Begin VB.Menu M_ACT_LIN5 
            Caption         =   "-"
         End
         Begin VB.Menu M_RPT_ACT_CLG_DAT 
            Caption         =   "&m Clearing Date Reports"
            Begin VB.Menu M_RPT_ACT_CLG_BBK 
               Caption         =   "&a Bank n Cash Book"
            End
            Begin VB.Menu M_RPT_ACT_CLG_DLT 
               Caption         =   "&b Debtors List"
            End
            Begin VB.Menu M_RPT_ACT_CLG_CRT 
               Caption         =   "&c Creditors List"
            End
            Begin VB.Menu M_RPT_ACT_CLG_CHQ_RCT 
               Caption         =   "&d Chqs Outstanding - Receipts"
            End
            Begin VB.Menu M_RPT_ACT_CLG_CHQ_PMT 
               Caption         =   "&e Chqs Outstanding - Payments"
            End
         End
         Begin VB.Menu M_RPT_ACT_HYP5 
            Caption         =   "-"
         End
         Begin VB.Menu M_RPT_ACT_FAC 
            Caption         =   "&n Final Accounts"
            Begin VB.Menu bls 
               Caption         =   "&a Balance Sheet - Detail"
            End
            Begin VB.Menu M_RPT_ACT_FAC_BLS 
               Caption         =   "&b Balance Sheet - Group"
            End
            Begin VB.Menu M_RPT_ACT_FAC_ICE 
               Caption         =   "&c Income n Expense"
            End
            Begin VB.Menu M_RPT_ACT_FAC_HYP_1 
               Caption         =   "-"
            End
            Begin VB.Menu M_RPT_ACT_FAC_TBL 
               Caption         =   "&d Trial Balance - Detail"
            End
            Begin VB.Menu M_RPT_ACT_FAC_GTBL 
               Caption         =   "&e Trial Balance - Group"
               Index           =   0
            End
         End
      End
      Begin VB.Menu MNU_RPT_HYP2 
         Caption         =   "-"
      End
      Begin VB.Menu m_rpt_oth 
         Caption         =   "&Regular"
         Enabled         =   0   'False
         Tag             =   "Regular"
         Begin VB.Menu M_RPT_OTH_DIP 
            Caption         =   "&a Discount Report"
            Index           =   0
         End
         Begin VB.Menu M_RPT_OTH_DTS 
            Caption         =   "&e Transactions List - Stock"
         End
         Begin VB.Menu M_RPT_OTH_DTF 
            Caption         =   "&f Transactions List - Finance"
         End
         Begin VB.Menu M_RPT_OTH_STR 
            Caption         =   "&g Status Report"
         End
         Begin VB.Menu M_RPT_OTH_HYP1 
            Caption         =   "-"
         End
         Begin VB.Menu M_RPT_OTH_CRC 
            Caption         =   "&h Variance  - Pur. Rate"
         End
         Begin VB.Menu M_RPT_OTH_DIF_pur 
            Caption         =   "&i Variance - Product - Purchase"
         End
         Begin VB.Menu M_RPT_OTH_SRC 
            Caption         =   "&j Variance - Sale Rate"
         End
         Begin VB.Menu M_RPT_OTH_DIF 
            Caption         =   "&k Variance - Product - Sales"
         End
         Begin VB.Menu M_RPT_OTH_HYP2 
            Caption         =   "-"
         End
         Begin VB.Menu M_RPT_OTH_RLT 
            Caption         =   "&l Retailers List - Tracking"
         End
         Begin VB.Menu M_RPT_OTH_DIS 
            Caption         =   "&m Retailer List - Discount"
         End
         Begin VB.Menu M_RPT_OTH_MLS 
            Caption         =   "&n Mailers"
         End
         Begin VB.Menu line 
            Caption         =   "-"
         End
         Begin VB.Menu M_RPT_OTH_COK 
            Caption         =   "&o Schemes && Prd'vty"
            Enabled         =   0   'False
            Tag             =   "MIS"
            Begin VB.Menu M_AR_COK_MIS 
               Caption         =   "&a Stock Summary"
               Begin VB.Menu M_RPT_OTH_COK_SSB 
                  Caption         =   "&a Batch - Volume"
               End
               Begin VB.Menu M_RPT_OTH_SSBR 
                  Caption         =   "&b Brand - Volume"
               End
               Begin VB.Menu M_RPT_OTH_SEG 
                  Caption         =   "&c Segment - Volume"
               End
            End
            Begin VB.Menu M_RPT_OTH_PNT 
               Caption         =   "&b Productivity Analysis"
            End
            Begin VB.Menu M_RPT_OTH_WSR 
               Caption         =   "&c FMCG Cal. Sales Report"
            End
            Begin VB.Menu M_RPT_OTH_CSEP 
               Caption         =   "-"
               Index           =   0
            End
            Begin VB.Menu M_RPT_OTH_CLM 
               Caption         =   "&d Scheme n Stock Report"
               Index           =   0
            End
         End
      End
      Begin VB.Menu MNU_RPT_HYP3 
         Caption         =   "-"
      End
      Begin VB.Menu M_RPT_RNG 
         Caption         =   "Range &Print"
         Enabled         =   0   'False
         Tag             =   "Range Print"
         Begin VB.Menu M_RPT_RNG_COR 
            Caption         =   "&Core"
            Enabled         =   0   'False
            Tag             =   "Core(RP)"
            Begin VB.Menu M_RPT_RNG_COR_SAL 
               Caption         =   "&a Sales"
            End
            Begin VB.Menu M_RPT_RNG_COR_HYP1 
               Caption         =   "-"
            End
            Begin VB.Menu M_RPT_RNG_COR_SSL 
               Caption         =   "&b Purchases"
            End
            Begin VB.Menu M_RPT_RNG_COR_OPS 
               Caption         =   "&c Opening Stock"
            End
            Begin VB.Menu M_RPT_RNG_COR_CRN 
               Caption         =   "&d Credit Notes"
            End
            Begin VB.Menu M_RPT_RNG_COR_DRN 
               Caption         =   "&e Debit Notes"
            End
            Begin VB.Menu M_RPT_RNG_COR_HYP3 
               Caption         =   "-"
            End
            Begin VB.Menu M_RPT_RNG_COR_STI 
               Caption         =   "&f Transfer Inward"
            End
            Begin VB.Menu M_RPT_RNG_COR_STO 
               Caption         =   "&g Transfer Outward"
            End
            Begin VB.Menu M_RPT_RNG_COR_VAN 
               Caption         =   "&h Van Challan"
            End
         End
         Begin VB.Menu M_RPT_RNG_ACT 
            Caption         =   "&Accounts"
            Tag             =   "Accounts(RP)"
            Begin VB.Menu M_RPT_RNG_ACT_RCT 
               Caption         =   "&a Receipts"
            End
            Begin VB.Menu M_RPT_RNG_ACT_PMT 
               Caption         =   "&b Payments"
            End
            Begin VB.Menu M_RPT_RNG_ACT_CRN 
               Caption         =   "&c Credit Notes"
            End
            Begin VB.Menu M_RPT_RNG_ACT_DRN 
               Caption         =   "&d Debit Notes"
            End
            Begin VB.Menu M_RPT_RNG_ACT_JVN 
               Caption         =   "&e Journal Entries"
            End
            Begin VB.Menu M_RPT_RNG_ACT_OPB 
               Caption         =   "&f Opening Balance"
            End
         End
      End
      Begin VB.Menu M_REP_HYP_5 
         Caption         =   "-"
      End
      Begin VB.Menu M_REP_MYPAGE 
         Caption         =   "&My Page"
         Begin VB.Menu M_RPT_MYPAGE_SQA 
            Caption         =   "&a Sales Quantity Analysis - ReadyStock"
            Index           =   0
         End
      End
      Begin VB.Menu M_REP_HYP_6 
         Caption         =   "-"
      End
      Begin VB.Menu M_REP_DIP 
         Caption         =   "&Direct Print"
         Enabled         =   0   'False
         Tag             =   "Direct Print"
      End
   End
   Begin VB.Menu M_HLP 
      Caption         =   "&Help"
      Begin VB.Menu M_HLP_EFF 
         Caption         =   "&About DAS"
         Shortcut        =   ^Z
      End
      Begin VB.Menu M_HLP_HLT 
         Caption         =   "&Help Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu M_SYS_CAL 
         Caption         =   "C&alculator"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu M_SYS 
      Caption         =   "&System"
      Begin VB.Menu SYS_SHT_STS 
         Caption         =   "Sh&utdownStatus"
      End
      Begin VB.Menu SYS_SET 
         Caption         =   "&Settings"
         Begin VB.Menu M_SYS_CCP 
            Caption         =   "&Company Switchover"
            Shortcut        =   ^J
         End
         Begin VB.Menu M_SYS_CLG 
            Caption         =   "Change &Login"
            Shortcut        =   ^L
         End
         Begin VB.Menu M_SYS_INI 
            Caption         =   "&Distributor Profile"
            Shortcut        =   ^T
         End
         Begin VB.Menu M_SYS_ACP 
            Caption         =   "&Financial Year"
            Index           =   0
            Shortcut        =   ^Y
         End
         Begin VB.Menu M_SYS_INV 
            Caption         =   "&Specify Defaults"
            Index           =   0
            Shortcut        =   ^F
         End
      End
      Begin VB.Menu SYS_DTB 
         Caption         =   "&Database"
         Enabled         =   0   'False
         Begin VB.Menu M_SYS_BKP 
            Caption         =   "&Validate Database"
            Shortcut        =   ^V
         End
         Begin VB.Menu M_SYS_IMP_MAS 
            Caption         =   "&Complete Backup"
            Index           =   0
            Shortcut        =   ^B
         End
      End
      Begin VB.Menu SYS_HYP_2 
         Caption         =   "-"
      End
      Begin VB.Menu SYS_MIS 
         Caption         =   "&Others"
         Visible         =   0   'False
         Begin VB.Menu M_SYS_BMP 
            Caption         =   "&Background Image"
            Index           =   0
         End
      End
   End
   Begin VB.Menu M_EXT 
      Caption         =   "E&xit"
      Begin VB.Menu M_EXT_WDP 
         Caption         =   "&Windows Desktop"
         Index           =   0
      End
   End
End
Attribute VB_Name = "M_MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private REC_POS As Integer
Private Sub bls_Click()
M_FRM_ASD.Tag = "Balance Sheet"
M_FRM_ASD.Show 1
End Sub
Private Sub dbt_agd_Click()
M_REP_SQY.Tag = "dbt_agd"
M_REP_SQY.Show 1
End Sub
Private Sub DebtorsAgeing_Click()
M_FRM_DBT.Tag = "Ageing"
M_FRM_DBT.Show 1
End Sub
Private Sub DIS_SCH_CPC_Click()
M_REP_SQY.Tag = "DSCPC"
    M_REP_SQY.Show 1
End Sub
Private Sub DIS_STK_REP_Click(i As Integer)
If i = 0 Then
    M_FRM_UPD.Tag = "DIS_STK_REP"
    M_FRM_UPD.Show 1
ElseIf i = 3 Then
    OpenReport CURDIR & "\STOCK.RP1", "Stock Breakup", 0, "", 6
ElseIf i = 4 Then
    M_FRM_DSN.Tag = "NSPD"
    M_FRM_DSN.Show 1
ElseIf i = 5 Then
    M_FRM_DSN.Tag = "NSPD_VAN"
    M_FRM_DSN.Show 1
Else
    M_FRM_DSN.Tag = "stkStmt"
    M_FRM_DSN.Show 1
End If
End Sub

Private Sub DLY_PRD_SUM_Click()
M_FRM_DSN.Tag = "DPS"
    M_FRM_DSN.Show
End Sub

Private Sub Image1_Click()
Shell "C:\PROGRAM FILES\INTERNET EXPLORER\iexplore.exe www.consoftsystems.com", vbMaximizedFocus
End Sub

Private Sub Label1_Click()
Shell "C:\PROGRAM FILES\INTERNET EXPLORER\iexplore.exe www.consoftsystems.com", vbMaximizedFocus
End Sub

Private Sub M_ACT_Click()
STS_BAR.Panels(1).Text = "DAS Version 4.0 Accounts Module: Payments, Receipts, Opening Balances & Journal Entries"
End Sub

Private Sub M_ACT_PMT_Click()
M_FRM_PMT.Show 1
End Sub

Private Sub M_ACT_RCT_Click()
M_FRM_RCT.Show 1
End Sub

Private Sub M_ACT_VCH_CRN_Click()
M_FRM_CRN.Show 1
End Sub

Private Sub M_ACT_VCH_DRN_Click()
M_FRM_DBN.Show 1
End Sub

Private Sub M_ACT_VCH_JVN_Click()
M_FRM_JRN.Show 1
End Sub

Private Sub M_ACT_VCH_OPB_Click()
M_FRM_OPN.Show 1
End Sub

Private Sub M_ANA_COM_RPT_Click()
M_RPT_COM.Show 1
End Sub

Private Sub M_ANA_PRG_CAT_Click()
    sprg = "Category"
    M_SAL_PRG.Show 1
End Sub

Private Sub M_ANA_PRG_LOC_Click()
sprg = "Location"
    M_SAL_PRG.Show 1
End Sub

Private Sub M_ANA_PRG_MAR_Click()
 sprg = "Market"
    M_SAL_PRG.Show 1
End Sub

Private Sub M_ANA_PRG_RET_Click()
sprg = "Retailer"
    M_SAL_PRG.Show 1
End Sub

Private Sub M_ANA_PRG_ROU_Click()
sprg = "Route"
    M_SAL_PRG.Show 1
End Sub

Private Sub M_ANA_PRG_RPT_Click()
'M_SAL_PRG.Show 1
End Sub

Private Sub M_ANA_PRG_SMN_Click()
sprg = "Salesman"
    M_SAL_PRG.Show 1
End Sub

Private Sub M_ANA_RET_LBD_Click()
M_REP_SQY.Tag = "LBD"
    M_REP_SQY.Show 1
End Sub

Private Sub M_ANA_RET_RRN_Click()
M_FRM_DSN.Tag = "RRN"
    M_FRM_DSN.Show 1
End Sub

Private Sub M_ANA_RET_RSN_Click()
M_FRM_DSN.Tag = "RSN"
    M_FRM_DSN.Show 1
End Sub

Private Sub M_ANA_RET_SKU_Click()
M_REP_STR.Tag = "TOP_SKU"
M_REP_STR.Show 1
End Sub

Private Sub M_ANA_SPC_RPT_Click()
'M_RPT_GTH.Show 1
'M_RPT_GWT.Show 1
End Sub

Private Sub M_AST_INW_Click()
M_FRM_AST_INW.Show 1
End Sub

Private Sub M_AST_OUT_Click()
M_FRM_AST_OUT.Show 1
End Sub

Private Sub M_COR_Click()
STS_BAR.Panels(1).Text = "DAS Version 4.0 Core Module: Sales, Purchases, Stocks & Ready Stocks"
End Sub
Private Sub M_COR_FRE_SAL_Click(i As Integer)
If i = 0 Then
    SETYEAR M_FRM_SAL
    M_FRM_SAL.Show 1
End If
End Sub
Private Sub M_COR_PSH_SAL_Click()
M_FRM_SAL_PSH.Show 1
End Sub
Private Sub M_COR_PUR_SSL_Click()
M_FRM_PUR.Show 1
End Sub
Private Sub M_COR_SAL_LDS_Click(i As Integer)
If i = 0 Then
    M_FRM_LDS.Show 1
'ElseIf I = 2 Then
'    M_FRM_VVR.Show 1
End If
End Sub
Private Sub M_COR_SAL_RSB_Click(i As Integer)
If i = 0 Then
    M_FRM_RSI1.Show 1
End If
End Sub
Private Sub M_COR_SHL_TRS_Click()
M_FRM_TRS.Show 1
End Sub
Private Sub M_COR_STK_ADJ_JRN_Click()
    M_FRM_ADJ.Show 1
End Sub
Private Sub M_COR_STK_COV_Click()
M_FRM_COV.Show 1
End Sub
Private Sub M_COR_STK_CRN_Click()
M_FRM_SRT.Show 1
End Sub
Private Sub M_COR_STK_DRN_Click()
M_FRM_PRT.Show 1
End Sub
Private Sub M_COR_STK_OPS_Click()
M_FRM_OPS.Show 1
End Sub
Private Sub M_COR_STK_STI_Click()
M_FRM_STI.Show 1
End Sub
Private Sub M_COR_STK_STO_Click(i As Integer)
If i = 0 Then
    M_FRM_STO.Show 1
End If
End Sub
Private Sub M_COR_TWO_REP_ESH_Click(Index As Integer)
If Index = 0 Then
'    M_REP_EMT.Tag = "SOH"
'    M_REP_EMT.Show 1
OpenReport CURDIR & "\SNN.RP1", "Shells SOH", 2, "", 6
ElseIf Index = 1 Then
'    M_REP_EMT.Tag = "ROLCOL"
'    M_REP_EMT.Show 1
    OpenReport CURDIR & "\SNN.RP1", "Shell ROL COL", 2, "(ACT->CST_TYP <> 'Van')", 6
ElseIf Index = 2 Then
'    Openreport CURDIR & "\SNN.RP1", "Party ROL COL", 2, "(ACT->CST_TYP <> 'Van')", 6
    M_REP_EMT.Tag = "PRC"
    M_REP_EMT.Show 1
ElseIf Index = 3 Then
'    Openreport CURDIR & "\SNN.RP1", "Party ROL COL Account", 2, "(ACT->CST_TYP <> 'Van')", 6
    M_REP_EMT.Tag = "PRCA"
    M_REP_EMT.Show 1
ElseIf Index = 4 Then
'    Openreport CURDIR & "\SNN.RP1", "Deposits Account", 2, "(ACT->CST_TYP <> 'Van')", 6
    M_REP_EMT.Tag = "SDA"
    M_REP_EMT.Show 1
ElseIf Index = 5 Then
    M_REP_EMT.Tag = "EMP_REG"
    M_REP_EMT.Show 1
    'M_FRM_DSN.Tag = "EMP_REG"
    'M_FRM_DSN.Show 1
ElseIf Index = 6 Then
    OpenReport CURDIR & "\SNN.RP1", "Shells SOH Net Load", 0, "", 6
End If
End Sub
Private Sub M_COR_UPS_Click()
M_MDI.MousePointer = vbHourglass
UPD_STK
M_MDI.MousePointer = vbDefault
'M_FRM_UPD.Show 1
End Sub
Private Sub M_COR_VAN_VVR_Click(Index As Integer)
If Index = 0 Then
    M_FRM_VVR.Show 1
End If
End Sub
Private Sub M_EXT_Click()
STS_BAR.Panels(1).Text = "DAS Version 4.0 Exit Module: Returns Control to Windows Desktop"
End Sub
Private Sub M_EXT_SHT_Click()
'If MsgBox("Close All Other Running Applications Before Shutdown", vbExclamation + vbOKCancel, "DAS Version 4.0 - Shutdown") = vbOK Then
'ExitWindowsEx 1, 0
'End If
End Sub
Private Sub M_EXT_WDP_Click(i As Integer)
If i = 0 Then
    M_FRM_EXT.Show 1
Else
    If MsgBox("Close All Other Running Applications Before Shutdown", vbExclamation + vbOKCancel, "DAS Version 4.0 - Shutdown") = vbOK Then
        ExitWindowsEx 1, 0
    End If
End If
End Sub
Private Sub M_FIL_ACT_Click()
M_FRM_ACT.Show 1
End Sub
Private Sub M_FIL_ASM_Click()
M_FRM_ASM.Show 1
End Sub
Private Sub M_FIL_ATT_Click()
M_FRM_ATT.Show 1
End Sub
Private Sub M_FIL_Click()
STS_BAR.Panels(1).Text = "DAS Version 4.0 File Module, Applicable to all functions of the system"
End Sub
Private Sub M_FIL_DAS_CRS_Click(Index As Integer)
    MsgBox ("Stock Update Process Required!"), vbInformation, "DAS Version 4.0"
    M_FRM_UPD.Show 1
    M_FRM_ASD.Tag = "DAS_CRS"
    M_FRM_ASD.Show 1
End Sub
Private Sub M_FIL_EXT_Click()
M_FRM_EXT.Show 1
End Sub
Private Sub M_FIL_GRP_Click()
M_FRM_GRP.Show 1
End Sub
Private Sub M_FIL_HYP1_Click(Index As Integer)
If Index = 1 Then
    M_FRM_ACT.Show 1
ElseIf Index = 2 Then
    M_FRM_GRP.Show 1
End If
End Sub
Private Sub M_FIL_HYP2_Click(Index As Integer)
If Index = 1 Then
    M_FRM_PGR.Show 1
ElseIf Index = 2 Then
    M_FRM_PRD.Show 1
End If
End Sub
Private Sub M_FIL_HYP5_Click(Index As Integer)
'If Index = 1 Then
'    M_FRM_RUT.Show 1
'ElseIf Index = 2 Then
'    M_FRM_SIC.Show 1
'ElseIf Index = 3 Then
'    M_FRM_INC.Show 1
'ElseIf Index = 4 Then
'    M_FRM_QPS.Show 1
'End If
    MsgBox ("Stock Update Process Required!"), vbInformation, "DAS Version 4.0"
    M_FRM_UPD.Show 1
    M_FRM_EXG.Show 1
End Sub
Private Sub M_FIL_LOC_Click()
M_FRM_LOC.Show 1
End Sub
Private Sub M_FIL_CAT_Click()
'M_FRM_CAT.Show 1
'If Index = 0 Then
 M_FRM_CAT.Show 1
'ElseIf Index = 1 Then
'    M_FRM_LOC.Show 1
'ElseIf Index = 2 Then
'    M_FRM_SMN.Show 1
'ElseIf Index = 3 Then
'    M_FRM_MAR.Show 1
'End If



End Sub
Private Sub M_FIL_MAR_Click()
M_FRM_MAR.Show 1
End Sub
Private Sub M_FIL_PGR_Click()
M_FRM_PGR.Show 1
End Sub
Private Sub M_FIL_PRD_Click()
M_FRM_PRD.Show 1
End Sub
Private Sub M_FIL_RAT_Click(Index As Integer)
'M_FRM_RAT.Show 1
End Sub
Private Sub M_FIL_RUT_Click()
'M_FRM_RUT.Show 1
End Sub
Private Sub M_FIL_SCH_Click(Index As Integer)
If Index = 0 Then
    M_FRM_SCH.Show 1
ElseIf Index = 1 Then   'HVO Schemes Form (to be developed)
    M_FRM_HVO.Show 1
ElseIf Index = 4 Then
    M_FRM_RAT.Show 1
ElseIf Index = 6 Then
    M_FRM_CAL.Show 1
ElseIf Index = 8 Then
    M_FRM_SIC.Show 1
ElseIf Index = 9 Then
    M_FRM_TGT.Show 1
ElseIf Index = 10 Then
    M_FRM_RUT.Show 1
'ElseIf Index = 14 Then
'    ETLForm.Show 1
'ElseIf Index = 15 Then
'    frmBarCode.Show 1
'ElseIf Index = 12 Then
'    frmSMS.Show 1
End If
End Sub
Private Sub M_FIL_SIC_Click(i As Integer)
If i = 0 Then
    M_FRM_SIC.Show 1
ElseIf i = 1 Then
    M_FRM_QPS.Show 1
End If
End Sub
Private Sub M_FIL_SMN_Click()
M_FRM_SMN.Show 1
End Sub
Private Sub m_fil_tgt_Click()
M_FRM_TGT.Show 1
End Sub
Private Sub M_FIL_CLD_Click()
M_FRM_CAL.Show 1
End Sub
Private Sub M_HLP_Click()
STS_BAR.Panels(1).Text = "DAS Version 4.0 Help System & Licensed Users Information"
End Sub
Private Sub M_HLP_EFF_Click()
M_FRM_ABT.Show 1
End Sub
Private Sub M_HLP_HLT_Click()
M_HELP_REG
End Sub
Private Sub m_mar_loc_Click()
M_FRM_PPQ.Tag = "PAR_LOC"
M_FRM_PPQ.Show 1
End Sub
Private Sub m_mir_cat_Click()
M_FRM_PPQ.Tag = "PAR_CAT"
M_FRM_PPQ.Show 1
End Sub
Private Sub M_MIR_CFR_Click()
M_REP_SQY.Tag = "CFR"
M_REP_SQY.Show 1
End Sub
Private Sub M_MIR_MAR_Click()
M_FRM_PPQ.Tag = "PAR_MAR"
M_FRM_PPQ.Show 1
End Sub
Private Sub M_MIR_PPD_SEG_Click()
M_REP_SQY.Tag = "PPD_SEG"
M_REP_SQY.Show 1
End Sub
Private Sub M_MIR_PPD_CAT_Click()
M_REP_SQY.Tag = "PPD_CAT"
M_REP_SQY.Show 1
End Sub

Private Sub M_MIR_ROU_Click(i As Integer)
If i = 0 Then
    M_FRM_PPQ.Tag = "PAR_ROU"
    M_FRM_PPQ.Show 1
End If
End Sub
Private Sub M_MIR_Click()
STS_BAR.Panels(1).Text = "DAS Version 4.0 Operations (Sales, Distributions & Purchases), Management Information Reports"
End Sub
Private Sub M_MIR_MBR_Click(i As Integer)
If i = 0 Then
    M_REP_SQY.Tag = "SAL_SGA"
    M_REP_SQY.Show 1
ElseIf i = 1 Then
    M_REP_SQY.Tag = "SAL_N14"
    M_REP_SQY.Show 1
ElseIf i = 2 Then
    M_REP_SQY.Tag = "SAL_RPR"
    M_REP_SQY.Show 1
ElseIf i = 3 Then
    M_FRM_ASD.Tag = "AGD_6WB"
    M_FRM_ASD.Show 1
ElseIf i = 4 Then
    M_REP_SQY.Tag = "SAL_BPS"
    M_REP_SQY.Show 1
ElseIf i = 5 Then
'    M_FRM_DBT.Tag = "ARR_COL"
'    M_FRM_DBT.Show 1
     M_REP_SQY.Tag = "ARR_COL"
     M_REP_SQY.Show 1
ElseIf i = 7 Then
    MBR_CLC_MOD
ElseIf i = 8 Then
    M_REP_STR.Tag = "STR"
    M_REP_STR.Show 1
ElseIf i = 9 Then
    M_REP_SQY.Tag = "MCFR"
    M_REP_SQY.Show 1
ElseIf i = 11 Then
    M_FRM_DSN.Tag = "SCH_QTY_MER"
    M_FRM_DSN.Show 1
ElseIf i = 12 Then
    M_FRM_DSN.Tag = "SCH_VAL_MER"
    M_FRM_DSN.Show 1
ElseIf i = 13 Then
    M_REP_SQY.Tag = "HVOS"
    M_REP_SQY.Show 1
ElseIf i = 14 Then
    M_REP_SQY.Tag = "HVOP"
    M_REP_SQY.Show 1
ElseIf i = 16 Then
    M_FRM_DSN.Tag = "EMP_MOV_RPT"
    M_FRM_DSN.Show 1
'    MsgBox ("Report Available as Update Soon"), vbInformation, "DAS Version 4.0"
ElseIf i = 17 Then
    MsgBox ("Report Available as Update Soon"), vbInformation, "DAS Version 4.0"
ElseIf i = 18 Then
    M_REP_SQY.Tag = "SMN_COM_PYB"
    M_REP_SQY.Show 1
ElseIf i = 19 Then
    M_REP_SQY.Tag = "ROU_GP_RPT"
    M_REP_SQY.Show 1
ElseIf i = 20 Then
    M_REP_SQY.Tag = "GDW_GP_RPT"
    M_REP_SQY.Show 1
ElseIf i = 22 Then
    M_REP_SQY.Tag = "FLRS"
    M_REP_SQY.Show 1
ElseIf i = 23 Then
    M_REP_SQY.Tag = "FLRD"
    M_REP_SQY.Show 1
    'MsgBox ("Report Available as Update Soon"), vbInformation, "DAS Version 4.0"
End If
End Sub
Private Sub MBR_CLC_MOD()
    M_MDI.MousePointer = vbHourglass
    Dim rs As New ADODB.Recordset
    Dim RS1 As New ADODB.Recordset
    con.Execute "DELETE FROM MBR"
    rs.Open "SELECT * FROM MBR", con, adOpenKeyset, adLockPessimistic
    For A = 4 To 12
        rs.AddNew "MTH_IDY", A
        rs.Update
    Next
    For A = 1 To 3
        rs.AddNew "MTH_IDY", A
        rs.Update
    Next
    rs.Requery
    For b = 1 To 12
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT SUM(NET_AMT),TXN_TYP FROM INL WHERE MONTH(INV_DAT)= " & b & "GROUP BY TXN_TYP", con, adOpenKeyset, adLockPessimistic
        While RS1.EOF = False
            rs.Find "MTH_IDY LIKE '" & b & "'"
            If rs.EOF = False Then
                If Trim(RS1(1)) = "SAL" Then
                    rs.Update "SAL_VAL", Val(rs("SAL_VAL") & "") + Val(RS1(0) & "")
                ElseIf Trim(RS1(1)) = "LSL" Then
                    rs.Update "LSL_VAL", Val(rs("LSL_VAL") & "") + Val(RS1(0) & "")
                ElseIf Trim(RS1(1)) = "PUR" Then
                    rs.Update "PUR_VAL", Val(rs("PUR_VAL") & "") + Val(RS1(0) & "")
                ElseIf Trim(RS1(1)) = "PRT" Then
                    rs.Update "PRT_VAL", Val(rs("PRT_VAL") & "") + Val(RS1(0) & "")
                ElseIf Trim(RS1(1)) = "SRN" Then
                    rs.Update "PRT_VAL", Val(rs("SRN_VAL") & "") + Val(RS1(0) & "")
                ElseIf Trim(RS1(1)) = "STI" Then
                    rs.Update "STI_VAL", Val(rs("STI_VAL") & "") + Val(RS1(0) & "")
                ElseIf Trim(RS1(1)) = "STO" Then
                    rs.Update "STO_VAL", Val(rs("STO_VAL") & "") + Val(RS1(0) & "")
                End If
            End If
            RS1.MoveNext
            rs.MoveFirst
        Wend
    Next
    M_MDI.MousePointer = vbDefault
    OpenReport CURDIR & "\MIR.RP1", "Monthly Business Record", 0, "", 6
End Sub
Private Sub M_MIR_NPD_Click()
M_FRM_SAQ.Tag = "NOP"
M_FRM_SAQ.Show 1
End Sub
Private Sub M_MIR_RAR_Click()
M_REP_SQY.Tag = "RAR"
M_REP_SQY.Show 1
End Sub
Private Sub M_MIR_RCG_Click()
OpenReport CURDIR & "\MIR.RP1", "Retailers Category Report", 2, "ACT->ACT_TYP='SAL'.OR.ACT->ACT_TYP='BOT'", 6
End Sub
Private Sub M_MIR_SOT_Click()
M_REP_SQY.Tag = "SOT"
M_REP_SQY.Show 1
End Sub
Private Sub M_MIR_SPD_Click()
M_REP_PQY.Tag = "SPD"
M_REP_PQY.Show 1
End Sub
Private Sub M_MIR_PPD_Click()
M_REP_PQY.Tag = "PPD"
M_REP_PQY.Show 1
End Sub
Private Sub M_MIR_PPV_Click(i As Integer)
If i = 0 Then
    M_REP_PQY.Tag = "PPV"
    M_REP_PQY.Show 1
End If
End Sub
Private Sub M_MIR_TOT_Click()
M_REP_SQY.Tag = "TOP"
M_REP_SQY.Show 1
End Sub
Private Sub M_MIR_YPV_Click()
M_REP_YER.Tag = "VAL"
M_REP_YER.Show 1
End Sub
Private Sub M_MIR_YQP_Click()
SAL_QRY = "INV->TXN_TYP='SAL'"
OpenReport CURDIR & "\MIR.RP1", "Yearly Products", 2, SAL_QRY, 6
End Sub
Private Sub M_MIR_YQC_Click()
SAL_QRY = "INV->TXN_TYP='SAL'"
OpenReport CURDIR & "\MIR.RP1", "Yearly Products Cases", 2, SAL_QRY, 6
End Sub
Private Sub M_MIR_YSR_Click()
SAL_QRY = "INL->TXN_TYP='SAL'"
OpenReport CURDIR & "\MIR.RP1", "Yearly Sales", 2, SAL_QRY, 6
End Sub

Private Sub M_MNU_SHL_INW_Click(i As Integer)
If i = 0 Then
    M_FRM_SIN.Show 1
End If
End Sub

Private Sub M_MNU_SHL_OUT_Click()
M_FRM_SOUT.Show 1
End Sub

Private Sub M_MNU_SHL_SNP_Click()
M_FRM_SHL.Show 1
End Sub

Private Sub M_REP_DIP_Click()
M_FRM_DRP.Show 1
End Sub
Private Sub M_RPT_ACT_BBK_Click()
M_FRM_AQY.Tag = "Bank Book"
M_FRM_AQY.Show 1
End Sub

Private Sub M_RPT_ACT_CLG_BBK_Click()
M_FRM_CLG_AQY.Tag = "CLG_DAY"
M_FRM_CLG_AQY.Show 1
End Sub

Private Sub M_RPT_ACT_CLG_CHQ_rct_Click()
M_FRM_DSN.Tag = "CLG_CHQ_RCT"
M_FRM_DSN.Show 1
End Sub
Private Sub M_RPT_ACT_CLG_CHQ_PMT_Click()
M_FRM_DSN.Tag = "CLG_CHQ_PMT"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_ACT_CLG_CRT_Click()
M_FRM_ASD.Tag = "CLG_CRT"
M_FRM_ASD.Show 1
End Sub

Private Sub M_RPT_ACT_CLG_DLT_Click()
M_REP_SQY.Tag = "CLG_DLT"
M_REP_SQY.Show 1
End Sub

Private Sub M_RPT_ACT_CNR_Click()
M_FRM_DSN.Tag = "CRN"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_ACT_CRL_Click()
M_FRM_ASD.Tag = "CRL"
M_FRM_ASD.Show 1
End Sub

Private Sub M_RPT_ACT_DLS_Click()
M_REP_SQY.Tag = "DRL"
M_REP_SQY.Show 1
End Sub

Private Sub M_RPT_ACT_DNR_Click()
M_FRM_DSN.Tag = "DRN"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_ACT_FAC_BLS_Click()
M_FRM_ASD.Tag = "Group Balance Sheet"
M_FRM_ASD.Show 1
End Sub

Private Sub M_RPT_ACT_FAC_ICE_Click()
M_FRM_ASD.Tag = "Income Statement"
M_FRM_ASD.Show 1
End Sub

Private Sub M_RPT_ACT_FAC_ROI_Click()
M_FRM_ASD.Tag = "ROI"
M_FRM_ASD.Show 1
End Sub

Private Sub M_RPT_ACT_FAC_TBL_Click()
M_FRM_ASD.Tag = "Trial Balance"
M_FRM_ASD.Show 1
End Sub
Private Sub M_RPT_ACT_FAC_GTBL_Click(i As Integer)
If i = 0 Then
    M_FRM_ASD.Tag = "Group Trial Balance"
    M_FRM_ASD.Show 1
ElseIf i = 2 Then

End If
End Sub

Private Sub M_RPT_ACT_GLD_Click()
M_FRM_AQY.Tag = "GLD"
M_FRM_AQY.Show
End Sub

Private Sub M_RPT_ACT_JVL_Click()
M_FRM_DSN.Tag = "JVL"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_ACT_OAL_Click()
M_FRM_ASD.Tag = "OPN"
M_FRM_ASD.Show 1
End Sub

Private Sub M_RPT_ACT_PVL_Click()
M_FRM_DSN.Tag = "PMT"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_ACT_RVL_Click()
M_FRM_DSN.Tag = "RCT"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_ACT_SMM_Click()
M_FRM_DSN.Tag = "Day Summary"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_Click()
STS_BAR.Panels(1).Text = "DAS Version 4.0 Reports Module, Applicable to all functions & modules of the system"
End Sub

Private Sub M_RPT_COR_PUR_BOK_Click()
M_FRM_DSN.Tag = "Purchase Book"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_PUR_PRD_Click()
M_PUR_PQY.Show 1
End Sub

Private Sub M_RPT_COR_PUR_REG_Click()
M_FRM_DSN.Tag = "Purchase Register"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_PUR_PRD_TAX_Click()
M_FRM_DSN.Tag = "Purchase Product Tax"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_PUR_PRD_TAX_TTL_Click()
M_FRM_DSN.Tag = "PUR_PRD_TAX_TTL"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_PUR_RET_Click()
M_FRM_PQY.Tag = "Purchase Return Report"
M_FRM_PQY.Show 1
End Sub

Private Sub M_RPT_COR_PUR_RPT_Click()
M_FRM_PQY.Tag = "Purchase Report"
M_FRM_PQY.Show 1
End Sub

Private Sub M_RPT_COR_PUR_SUM_Click()
M_FRM_PQY.Tag = "Purchase Summary"
M_FRM_PQY.Show 1
End Sub
Private Sub M_RPT_COR_PUR_TID_RPT_Click()
M_FRM_PQY.Tag = "Transit Report"
M_FRM_PQY.Show 1
End Sub

Private Sub M_RPT_COR_RST_RSR_Click()
M_FRM_DSN.Tag = "Ready Stocks Report (Issues)"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_RST_SIT_Click()
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim TPD As New ADODB.Recordset
con.Execute "DELETE FROM TPD"
TPD.Open "SELECT * FROM TPD", con, adOpenKeyset, adLockPessimistic
rs.Open "SELECT INV_IDY FROM INL WHERE TXN_TYP LIKE 'RSI'", con, adOpenKeyset, adLockPessimistic
For A = 1 To rs.RecordCount
    If RS1.State = 1 Then RS1.Close
    RS1.Open "SELECT * FROM INL WHERE ORD_IDY LIKE '" & Trim(rs("INV_IDY")) & "' AND TXN_TYP LIKE 'RSR'", con, adOpenKeyset, adLockPessimistic
    If RS1.RecordCount = 0 Then
        If RS1.State = 1 Then RS1.Close
        RS1.Open "SELECT * FROM INV,PRD WHERE PRD.PRD_IDY LIKE INV.PRD_IDY AND INV_IDY LIKE '" & Trim(rs("INV_IDY")) & "' AND TXN_TYP LIKE 'RSI'", con, adOpenKeyset, adLockPessimistic
        For b = 1 To RS1.RecordCount
            TPD.Find "PRD_IDY LIKE '" & Trim(RS1("PRD_IDY")) & "'"
            If TPD.EOF = False Then
                TPD("PRD_QTY") = STM(MTS(Val(TPD("PRD_QTY") & ""), RS1("SUB_UNT")) + MTS((RS1("PRD_QTY") & ""), RS1("SUB_UNT")) - MTS(Val(RS1("RET_QTY") & ""), RS1("SUB_UNT")), RS1("SUB_UNT"))
                TPD.Update
            Else
                TPD.AddNew
                TPD("PRD_IDY") = RS1("PRD_IDY")
                TPD("PRD_QTY") = STM(MTS((RS1("PRD_QTY") & ""), RS1("SUB_UNT")) - MTS(Val(RS1("RET_QTY") & ""), RS1("SUB_UNT")), RS1("SUB_UNT"))
                TPD("PRD_VAL") = 0
                TPD("PRD_UNR") = 0
                TPD.Update
            End If
            TPD.Requery
            RS1.MoveNext
        Next
    End If
    rs.MoveNext
Next
OpenReport CURDIR & "\REPORT.RP1", "Stock in Transit", 0, "", 6
End Sub

Private Sub M_RPT_COR_RST_SRP_Click()
M_REP_SQY.Tag = "VAN"
M_REP_SQY.Show 1
End Sub

Private Sub M_RPT_COR_SAL_PRD_INR_Click()
M_REP_PQY.Tag = "Invoice Returns"
M_REP_PQY.Show 1
End Sub

Private Sub M_RPT_COR_SAL_PRD_SCH_Click()
M_REP_PQY.Tag = "Sales Scheme"
M_REP_PQY.Show 1
End Sub

Private Sub M_RPT_COR_SAL_PRD_SCH_DAT_Click()
M_REP_PQY.Tag = "Sales Scheme Date"
M_REP_PQY.Show 1
End Sub

Private Sub M_RPT_COR_SAL_PRD_SCH_VAL_Click()
M_REP_PQY.Tag = "Scheme Value"
M_REP_PQY.Show 1
End Sub

Private Sub M_RPT_COR_SAL_PRD_TYP_Click()
M_REP_PQY.Tag = "Stock Type"
M_REP_PQY.Show 1
End Sub

Private Sub M_RPT_COR_SAL_PRD_TYP_GRP_Click()
M_REP_PQY.Tag = "Sales Product Type"
M_REP_PQY.Show 1
End Sub

Private Sub M_RPT_COR_SAL_RET_PRD_Click()
M_FRM_DSN.Tag = "Sales Return Product"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_SAL_TPT_TAX_Click()
M_FRM_DSN.Tag = "Sales Tax Product"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_SAL_TPT_TAX_TTL_Click()
M_FRM_DSN.Tag = "SAL_PRD_TAX_TTL"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_STK_ANL_AGD_VAL_Click()
'M_FRM_ASD.Tag = "AGD_VAL"
'M_FRM_ASD.Show
End Sub

Private Sub M_RPT_COR_STK_ANL_AGD_VAL_MFG_Click()
'M_FRM_ASD.Tag = "AGD_MFG"
'M_FRM_ASD.Show
End Sub

Private Sub M_RPT_COR_STK_ANL_PRD_Click()
'M_FRM_DSN.Tag = "PRD"
'M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_STK_ANL_SUN_Click()
'M_FRM_DSN.Tag = "SUN"
'M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_STK_RGL_NOD_INV_Click()
'M_FRM_DSN.Tag = "Inventory"
'M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_STK_RGL_STI_PRD_Click()
'M_FRM_DSN.Tag = "Inward Product"
'M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_STK_RGL_STL_DBR_Click()
'M_FRM_PPQ.Tag = "DMG_BRK"
'M_FRM_PPQ.Show 1
End Sub

Private Sub M_RPT_COR_STK_RGL_STL_DBR_SEG_Click()
'M_FRM_PPQ.Tag = "DMG_BRK_SEG"
'M_FRM_PPQ.Show 1
End Sub

Private Sub M_RPT_COR_STK_RGL_STO_PRD_Click()
M_FRM_DSN.Tag = "Outward Product"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_MYPAGE_SQA_Click(i As Integer)
If i = 0 Then
    M_FRM_ASD.Tag = "MyPage_SQA"
    M_FRM_ASD.Show 1
ElseIf i = 1 Then
    M_FRM_ASD.Tag = "MyPage_SVA"
    M_FRM_ASD.Show 1
ElseIf i = 2 Then
    M_FRM_ASD.Tag = "MyPage_FUA"
    M_FRM_ASD.Show 1
ElseIf i = 3 Then
    M_FRM_ASD.Tag = "MyPage_VUT"
    M_FRM_ASD.Show 1
ElseIf i = 4 Then
    M_FRM_ASD.Tag = "MyPage_VUV"
    M_FRM_ASD.Show 1
ElseIf i = 5 Then
    M_FRM_ASD.Tag = "MyPage_STS"
    M_FRM_ASD.Show 1
End If
End Sub

Private Sub M_RPT_OTH_CLM_Click(i As Integer)
If i = 0 Then
    M_FRM_DSN.Tag = "Claims Report"
    M_FRM_DSN.Show 1
ElseIf i = 1 Then
    M_FRM_DSN.Tag = "Claims Report SKU" '"Claims Report Brand"
    M_FRM_DSN.Show 1
ElseIf i = 2 Then
    M_FRM_DSN.Tag = "Claims Report Brand" '"Claims Report Segment"
    M_FRM_DSN.Show 1
ElseIf i = 3 Then
    M_FRM_DSN.Tag = "Claims Report Segment" '"Claims Report SKU"
    M_FRM_DSN.Show 1
End If
End Sub

Private Sub M_RPT_OTH_COK_SSB_Click()
M_FRM_DSN.Tag = "COK_STK_BAT"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_OTH_CSEP_Click(Index As Integer)
If Index = 1 Then
    SAL_QRY = "inV->inv_dat >= ctod('" & Format(DTPicker1, "MM/DD/YYYY") & "').and.inv->inv_dat <= ctod('" & Format(DTPicker2, "MM/DD/YYYY") & "')"
    OpenReport CURDIR & "\CMP_RPT.RP1", "SKU Wise Route Sales", 0, Trim(SAL_QRY), 6
End If
End Sub

Private Sub M_RPT_OTH_DIP_Click(i As Integer)
If i = 0 Then
    M_REP_SQY.Tag = "DPT"
    M_REP_SQY.Show 1
ElseIf i = 1 Then
    M_REP_PQY.Tag = "RET_DIS"
    M_REP_PQY.Show 1
ElseIf i = 2 Then
    M_REP_PQY.Tag = "SEG_DIS"
    M_REP_PQY.Show 1
ElseIf i = 3 Then
    M_REP_PQY.Tag = "SUP_DIS"
    M_REP_PQY.Show 1
End If
End Sub

Private Sub M_RPT_OTH_VSC_Click()
M_FRM_RQY.Show 1
End Sub

Private Sub M_RPT_COR_RST_SRP_INV_Click()
M_FRM_RQY.Show 1
End Sub

Private Sub M_RPT_COR_RST_VRS_Click()
M_FRM_DSN.Tag = "Van Challan Returns"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_SAL_PRD_Click()
M_REP_PQY.Show 1
End Sub

Private Sub M_RPT_COR_SAL_PRD_FPR_Click()
M_REP_PQY.Tag = "Free"
M_REP_PQY.Show 1
End Sub
Private Sub M_RPT_OTH_FRS_Click()
M_REP_PQY.Tag = "Free"
M_REP_PQY.Show 1
End Sub
Private Sub M_RPT_COR_SAL_REG_Click()
M_FRM_DSN.Tag = "Sales Register"
M_FRM_DSN.Show 1
End Sub
Private Sub M_RPT_COR_SAL_RET_Click()
M_FRM_DSN.Tag = "Sales Return Report"
M_FRM_DSN.Show 1
End Sub
Private Sub M_RPT_COR_SAL_RPT_Click(i As Integer)
If i = 0 Then
    M_REP_SQY.Tag = "Sales Report"
    M_REP_SQY.Show 1
ElseIf i = 1 Then
    M_REP_SQY.Tag = "Sales Summary(Customer)"
    M_REP_SQY.Show 1
ElseIf i = 2 Then
    M_REP_SQY.Tag = "Sales Summary"
    M_REP_SQY.Show 1
ElseIf i = 3 Then
    M_FRM_DSN.Tag = "Sales Register"
    M_FRM_DSN.Show 1
ElseIf i = 4 Then
    M_REP_PQY.Show 1
ElseIf i = 5 Then
    M_FRM_DSN.Tag = "Sales Return Report"
    M_FRM_DSN.Show 1
ElseIf i = 6 Then
    M_FRM_DSN.Tag = "Sales Return Product"
    M_FRM_DSN.Show 1
ElseIf i = 7 Then
    M_REP_PQY.Tag = "Stock Type"
    M_REP_PQY.Show 1
ElseIf i = 8 Then
    M_REP_PQY.Tag = "Sales Product Type"
    M_REP_PQY.Show 1
ElseIf i = 9 Then
    M_REP_PQY.Tag = "Invoice Returns"
    M_REP_PQY.Show 1
ElseIf i = 10 Then
    M_REP_PQY.Tag = "Sales Scheme"
    M_REP_PQY.Show 1
ElseIf i = 11 Then
    M_REP_PQY.Tag = "Sales Scheme Date"
    M_REP_PQY.Show 1
ElseIf i = 12 Then
    M_REP_PQY.Tag = "Scheme Value"
    M_REP_PQY.Show 1
ElseIf i = 13 Then
    M_FRM_DSN.Tag = "Tax Report"
    M_FRM_DSN.Show 1
ElseIf i = 14 Then
    M_FRM_DSN.Tag = "Sales Tax Product"
    M_FRM_DSN.Show 1
ElseIf i = 15 Then
    M_FRM_DSN.Tag = "SAL_PRD_TAX_TTL"
    M_FRM_DSN.Show 1
ElseIf i = 16 Then
    M_REP_SQY.Tag = "Load Reports"
    M_REP_SQY.Show 1
ElseIf i = 17 Then
    M_FRM_DSN.Tag = "SAL_VAT"
    M_FRM_DSN.Show 1
ElseIf i = 18 Then
    M_FRM_DSN.Tag = "PUR_VAT"
    M_FRM_DSN.Show 1
ElseIf i = 19 Then
    M_FRM_DSN.Tag = "VAT_CAL"
    M_FRM_DSN.Show 1
ElseIf i = 20 Then
    M_FRM_DSN.Tag = "UOM_SAL"
    M_FRM_DSN.Show 1
ElseIf i = 21 Then
    M_FRM_DSN.Tag = "SCH_REF"
    M_FRM_DSN.Show 1
ElseIf i = 22 Then
    M_REP_SQY.Tag = "SSU_RCH"
    M_REP_SQY.Show 1
ElseIf i = 23 Then
    M_REP_SQY.Tag = "SSU_RLO"
    M_REP_SQY.Show 1
ElseIf i = 24 Then
    M_REP_SQY.Tag = "SSU_RMA"
    M_REP_SQY.Show 1
ElseIf i = 25 Then
    M_REP_SQY.Tag = "SSU_RRO"
    M_REP_SQY.Show 1
ElseIf i = 26 Then
    M_REP_SQY.Tag = "SSU_RSM"
    M_REP_SQY.Show 1
ElseIf i = 27 Then
    M_REP_PQY.Tag = "SRP_SML"
    M_REP_PQY.Show 1
ElseIf i = 28 Then
    M_REP_SQY.Tag = "SRP_DAC"
    M_REP_SQY.Show 1
ElseIf i = 29 Then
    M_REP_PQY.Tag = "SRP_SMY"
    M_REP_PQY.Show 1
ElseIf i = 30 Then
    M_REP_PQY.Tag = "BRD_SMY"
    M_REP_PQY.Show 1
ElseIf i = 31 Then
    M_REP_PQY.Tag = "PCK_SMY"
    M_REP_PQY.Show 1
ElseIf i = 32 Then
    M_REP_PQY.Tag = "PCT_SMY"
    M_REP_PQY.Show 1
ElseIf i = 33 Then
    M_FRM_QUK_RST.Tag = "XL_SAL"
    M_FRM_QUK_RST.Show 1
ElseIf i = 34 Then
    M_FRM_QUK_RST.Tag = "XL_VAN"
    M_FRM_QUK_RST.Show 1
End If
End Sub

Private Sub M_RPT_COR_SAL_RPT_CUS_Click()
M_REP_SQY.Tag = "Sales Summary(Customer)"
M_REP_SQY.Show 1
End Sub

Private Sub M_RPT_COR_SAL_SUM_Click()
M_REP_SQY.Tag = "Sales Summary"
M_REP_SQY.Show 1
End Sub

Private Sub M_RPT_COR_SAL_TPT_Click()
M_FRM_DSN.Tag = "Tax Report"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_STK_ANL_AGD_Click(i As Integer)
If i = 0 Then
    M_FRM_ASD.Tag = "AGD"
    M_FRM_ASD.Show
ElseIf i = 1 Then
    M_FRM_ASD.Tag = "AGD_CAS"
    M_FRM_ASD.Show
ElseIf i = 2 Then
    M_FRM_ASD.Tag = "AGD_VAL"
    M_FRM_ASD.Show
ElseIf i = 3 Then
    M_FRM_ASD.Tag = "AGD_MFG"
    M_FRM_ASD.Show
ElseIf i = 4 Then
    M_FRM_DSN.Tag = "SSM"
    M_FRM_DSN.Show 1
ElseIf i = 5 Then
    M_FRM_DSN.Tag = "SUN"
    M_FRM_DSN.Show 1
ElseIf i = 6 Then
    M_FRM_DSN.Tag = "PRD"
    M_FRM_DSN.Show 1
ElseIf i = 7 Then
'    M_FRM_ASD.Tag = "AGD"
'    M_FRM_ASD.Show
ElseIf i = 8 Then
    M_FRM_DSN.Tag = "Opening Stocks"
    M_FRM_DSN.Show 1
ElseIf i = 9 Then
    M_FRM_DSN.Tag = "Stock Issues"
    M_FRM_DSN.Show 1
ElseIf i = 10 Then
    M_FRM_DSN.Tag = "Inward Product"
    M_FRM_DSN.Show 1
ElseIf i = 11 Then
    M_FRM_DSN.Tag = "Stock Receipts"
    M_FRM_DSN.Show 1
ElseIf i = 12 Then
    M_FRM_DSN.Tag = "Outward Product"
    M_FRM_DSN.Show 1
ElseIf i = 13 Then
    M_FRM_ASD.Tag = "AGD"
    M_FRM_ASD.Show
ElseIf i = 14 Then
    M_FRM_ASD.Tag = "AGD"
    M_FRM_ASD.Show
ElseIf i = 15 Then
    M_FRM_ASD.Tag = "AGD"
    M_FRM_ASD.Show
ElseIf i = 16 Then
    M_FRM_ASD.Tag = "AGD"
    M_FRM_ASD.Show
End If
End Sub

Private Sub M_RPT_COR_STK_ANL_BST_Click()
M_FRM_PPQ.Tag = "Balance Stock"
M_FRM_PPQ.Show 1
End Sub

Private Sub M_RPT_COR_STK_ANL_GRB_Click()
UPD_STK
OpenReport CURDIR & "\stock.RP1", "Group Balance", 0, "", 6
End Sub

Private Sub M_RPT_COR_STK_ANL_MVR_Click(i As Integer)
If i = 0 Then
    M_REP_SQY.Tag = "SMR"
    M_REP_SQY.Show 1
End If
End Sub

Private Sub M_RPT_COR_STK_ANL_ROL_Click()
UPD_STK
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
con.Execute "DELETE FROM TPD"
For A = 1 To rs.RecordCount
    con.Execute "INSERT INTO TPD VALUES('" & Trim(rs("PRD_IDY")) & "'," & CONQTY(rs("REO_LEV"), rs("PRD_IDY")) & ",0,0)"
    rs.MoveNext
Next
OpenReport CURDIR & "\stock.RP1", "Reorder level", 0, "", 6
End Sub

Private Sub M_RPT_COR_STK_ANL_SFL_Click()
UPD_STK
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM PRD", con, adOpenKeyset, adLockPessimistic
con.Execute "DELETE FROM TPD"
For A = 1 To rs.RecordCount
    con.Execute "INSERT INTO TPD VALUES('" & Trim(rs("PRD_IDY")) & "'," & CONQTY(rs("SAF_LEV"), rs("PRD_IDY")) & ",0,0)"
    rs.MoveNext
Next
OpenReport CURDIR & "\stock.RP1", "Safety Level", 0, "", 6
End Sub

Private Sub M_RPT_COR_STK_ANL_SGV_Click()
UPD_STK
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If REC_SET1.State = 1 Then REC_SET1.Close
If REC_SET.State = 1 Then REC_SET.Close
con.Execute "delete from tpd"
REC_SET.Open "select * from prd", con, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then
    If REC_SET1.State = 1 Then REC_SET.Close
    REC_SET1.Open "select * from tpd", con, adOpenKeyset, adLockPessimistic
    For A = 1 To REC_SET.RecordCount
        REC_SET1.AddNew
        REC_SET1(0) = Trim(REC_SET("prd_idy") & "")
        REC_SET1(1) = QTY(REC_SET("cur_lev") & "")
        REC_SET1(3) = Val(REC_SET("bas_rat") & "")
        REC_SET1(2) = RATCAL(QTY(REC_SET("cur_lev") & ""), Val(REC_SET("bas_rat") & ""), Val(REC_SET("Sub_unt") & ""))
        REC_SET1.Update
        REC_SET.MoveNext
    Next
End If
OpenReport CURDIR & "\Stock.RP1", "Stock Value (Cat)", 0, "vch->txn_idy='" & Trim(INC) & "'", 6
End Sub

Private Sub M_RPT_COR_STK_ANL_SSM_Click()
'M_FRM_DSN.Tag = "SSM"
'M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_STK_ANL_VLR_Click()
UPD_STK
Dim REC_SET As New ADODB.Recordset
Dim REC_SET1 As New ADODB.Recordset
If REC_SET1.State = 1 Then REC_SET1.Close
If REC_SET.State = 1 Then REC_SET.Close
con.Execute "delete from tpd"
REC_SET.Open "select * from prd", con, adOpenKeyset, adLockPessimistic
If REC_SET.RecordCount > 0 Then
    If REC_SET1.State = 1 Then REC_SET.Close
    REC_SET1.Open "select * from tpd", con, adOpenKeyset, adLockPessimistic
    For A = 1 To REC_SET.RecordCount
        REC_SET1.AddNew
        REC_SET1(0) = Trim(REC_SET("prd_idy") & "")
        REC_SET1(1) = QTY(REC_SET("cur_lev") & "")
        REC_SET1(3) = Val(REC_SET("bas_rat") & "")
        REC_SET1(2) = RATCAL(QTY(REC_SET("cur_lev") & ""), Val(REC_SET("bas_rat") & ""), Val(REC_SET("Sub_unt") & ""))
        REC_SET1.Update
        REC_SET.MoveNext
    Next
End If
OpenReport CURDIR & "\Stock.RP1", "Stock Value", 0, "vch->txn_idy='" & Trim(INC) & "'", 6
End Sub


Private Sub M_RPT_COR_STK_RGL_OPS_Click()
'M_FRM_DSN.Tag = "Opening Stocks"
'M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_STK_RGL_STI_Click()
'M_FRM_DSN.Tag = "Stock Issues"
'M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_COR_STK_RGL_STL_Click(i As Integer)
If i = 0 Then
    M_FRM_PPQ.Tag = "Stock Ledger"
    M_FRM_PPQ.Show 1
ElseIf i = 1 Then
    M_FRM_PPQ.Tag = "Stock Ledger Test"
    M_FRM_PPQ.Show 1
ElseIf i = 2 Then
    M_FRM_PPQ.Tag = "DMG_BRK"
    M_FRM_PPQ.Show 1
ElseIf i = 3 Then
    M_FRM_PPQ.Tag = "DMG_BRK_SEG"
    M_FRM_PPQ.Show 1
ElseIf i = 4 Then
    M_FRM_DSN.Tag = "Inventory"
    M_FRM_DSN.Show 1
End If
End Sub

Private Sub M_RPT_COR_STK_RGL_STO_Click()
'M_FRM_DSN.Tag = "Stock Receipts"
'M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_OTH_A9S_Click()
M_FRM_DSN.Tag = "A9"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_OTH_CRC_Click()
con.Execute "Update dir where dir_idy like 'REP' and dir_idx like 'STT' SET DIR_MSG ='Purchase Rate Changes'"
Dim rs As New ADODB.Recordset
rs.Open "SELECT PRD_IDY,PRD_RAT,INV_DAT FROM INV WHERE TXN_TYP LIKE 'PUR' GROUP BY PRD_RAT", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    con.Execute "DELETE FROM RTP"
    For A = 1 To rs.RecordCount
        con.Execute "INSERT INTO RTP VALUES('" & Trim(rs(0)) & "'," & Val(rs(1)) & ",CTOD('" & rs(2) & "'))"
        rs.MoveNext
    Next
End If
OpenReport CURDIR & "\REPORT.RP1", "Rates Fluctuation", 0, "", 6
End Sub

Private Sub M_RPT_OTH_DIF_Click()
M_REP_SQY.Tag = "RAT"
M_REP_SQY.Show 1
End Sub

Private Sub M_RPT_OTH_DIF_pur_Click()
M_FRM_PQY.Tag = "Variance"
M_FRM_PQY.Show 1
End Sub

Private Sub M_RPT_OTH_DIS_Click()
M_REP_RET.Tag = "DIS"
M_REP_RET.Show 1
End Sub

Private Sub M_RPT_OTH_DTF_Click()
M_FRM_DSN.Tag = "Daily(Finance)"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_OTH_DTS_Click()
M_FRM_DSN.Tag = "Daily(Stock)"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_OTH_LSS_Click()
RNG_PRT_FRM.Tag = "LSS"
RNG_PRT_FRM.Caption = "Loading Sheet"
RNG_PRT_FRM.Show 1
End Sub
Private Sub M_RPT_OTH_SRD_Click()
M_REP_SQY.Tag = "Sales Report"
M_REP_SQY.Show 1
End Sub

Private Sub M_RPT_OTH_MLS_Click()
M_REP_RET.Tag = "MLS"
M_REP_RET.Show 1
End Sub

Private Sub M_RPT_OTH_PNT_Click()
M_REP_TGT.Tag = "PNT"
M_REP_TGT.Show 1
End Sub

Private Sub M_RPT_OTH_RLT_Click()
M_REP_RET.Tag = "LST"
M_REP_RET.Show 1
End Sub

Private Sub M_RPT_OTH_SEG_Click()
M_FRM_DSN.Tag = "COK_STK_SEG"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_OTH_SRC_Click()
con.Execute "Update dir where dir_idy like 'REP' and dir_idx like 'STT' SET DIR_MSG ='Sales Rate Changes'"
Dim rs As New ADODB.Recordset
rs.Open "SELECT PRD_IDY,PRD_RAT,INV_DAT FROM INV WHERE TXN_TYP in('LSL','SAL') GROUP BY PRD_RAT", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    con.Execute "DELETE FROM RTP"
    For A = 1 To rs.RecordCount
        con.Execute "INSERT INTO RTP VALUES('" & Trim(rs(0)) & "'," & Val(rs(1)) & ",CTOD('" & rs(2) & "'))"
        rs.MoveNext
    Next
End If
OpenReport CURDIR & "\REPORT.RP1", "Rates Fluctuation", 0, "", 6
End Sub

Private Sub M_RPT_OTH_SSBR_Click()
M_FRM_DSN.Tag = "COK_STK_BRD"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_OTH_STR_Click()
M_FRM_DSN.Tag = "STS"
M_FRM_DSN.Show 1
End Sub

Private Sub M_RPT_OTH_WSR_Click()
M_FRM_WSR.Show 1
End Sub

Private Sub M_RPT_RNG_ACT_CRN_Click()
RNG_PRT_FRM.Tag = "CRN"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_ACT_DRN_Click()
RNG_PRT_FRM.Tag = "DRN"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_ACT_JVN_Click()
RNG_PRT_FRM.Tag = "JRN"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_ACT_OPB_Click()
RNG_PRT_FRM.Tag = "OPA"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_ACT_PMT_Click()
RNG_PRT_FRM.Tag = "PMT"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_ACT_RCT_Click()
RNG_PRT_FRM.Tag = "RCT"
RNG_PRT_FRM.Show 1
End Sub


Private Sub M_RPT_RNG_COR_CRN_Click()
RNG_PRT_FRM.Tag = "SRN"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_COR_DRN_Click()
RNG_PRT_FRM.Tag = "PRT"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_COR_OPS_Click()
RNG_PRT_FRM.Tag = "OPS"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_COR_RSK_PCS_Click()
RNG_PRT_FRM.Tag = "PCS"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_COR_RSK_UNT_Click()
RNG_PRT_FRM.Tag = "UNT"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_COR_SAL_Click()
RNG_PRT_FRM.Tag = "SAL"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_COR_SSL_Click()
RNG_PRT_FRM.Tag = "PUR"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_COR_STI_Click()
RNG_PRT_FRM.Tag = "STI"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_COR_STO_Click()
RNG_PRT_FRM.Tag = "STO"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_COR_VAN_Click()
RNG_PRT_FRM.Tag = "RSI"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_RPT_RNG_COR_VSL_Click()
RNG_PRT_FRM.Tag = "VAN"
RNG_PRT_FRM.Show 1
End Sub

Private Sub M_SYS_ACP_Click(i As Integer)
If i = 0 Then
    YER_FRM.Show 1
ElseIf i = 1 Then
    Shell App.Path & "\Rapidrun.exe"
End If
End Sub

Private Sub M_SYS_BKP_Click()
checkfiles.Show 1
'BAC_UP_FRM.Show 1
End Sub

Private Sub M_SYS_BMP_Click(i As Integer)
On Error GoTo ErrHandler
If i = 0 Then
    If M_SYS_BMP(0).Checked = True Then
        M_SYS_BMP(0).Checked = False
        con.Execute "update dir set dir_msg='' where dir_idy like 'BMP'"
        Picture1.Picture = LoadPicture("")
    Else
        M_SYS_BMP(0).Checked = True
        CommonDialog1.filter = "All Files (*.*)|*.*|Bitmap Files (*.bmp)|*.bmp"
        CommonDialog1.DialogTitle = "DAS Version 4.0 Select Bitmap Image"
        CommonDialog1.FilterIndex = 2
        CommonDialog1.ShowOpen
        Picture1.Picture = LoadPicture(CommonDialog1.fileName)
        con.Execute "UPDATE DIR WHERE DIR_IDY LIKE 'BMP' AND DIR_IDX LIKE 'IDY' SET DIR_MSG ='" & Trim(CommonDialog1.fileName) & "'"
        ChDir CURDIR
    End If
    Exit Sub
ElseIf i = 1 Then
    FUN_KEY.Show 1
ElseIf i = 2 Then
    SYS_PSW.Tag = "HHT"
    SYS_PSW.Show 1
ElseIf i = 3 Then
'    m_FRM_HUP.Show 1
ElseIf i = 4 Then
'    M_FRM_EVP.Show 1
ElseIf i = 5 Then
'    M_FRM_MIBS.Show 1
End If
ErrHandler:

End Sub

Private Sub M_SYS_CAL_Click()
If osversion = 1 Then Shell "c:\windows\calc.exe", vbMaximizedFocus
If osversion = 2 Then Shell "c:\winnt\system32\calc.exe", vbMaximizedFocus
End Sub

Private Sub M_SYS_CBB_Click()
con.Execute "update dir set dir_msg='' where dir_idy like 'BMP'"
End Sub

Private Sub M_SYS_CCP_Click()
Dim DIR_RS As New ADODB.Recordset
If DIR_RS.State = 1 Then DIR_RS.Close
DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'IND'"), con, adOpenKeyset, adLockPessimistic
If DIR_RS.RecordCount > 0 Then
    DIR_RS.MoveFirst
    DIR_RS("DIR_MSG") = "PROPER SHUTDOWN"
    DIR_RS("DIR_DAT") = Date
    DIR_RS.Update
End If
DEL_COM.Show 1
End Sub

Private Sub M_SYS_CLG_Click()
EFF_LOG.Tag = "CHANGE"
EFF_LOG.Show 1
End Sub

Private Sub M_SYS_Click()
STS_BAR.Panels(1).Text = "DAS Version 4.0 System Module: File Management, Backup & Setup features"
End Sub

Private Sub M_SYS_FKD_Click()

End Sub

Private Sub M_SYS_IMP_MAS_Click(i As Integer)
On Error Resume Next
If i = 0 Then
    BAC_UP_FRM.Show 1
'    M_FRM_IFD.Tag = "Masters"
'    M_FRM_IFD.Show 1
ElseIf i = 1 Then
    M_FRM_EFD.Show 1
ElseIf i = 2 Then
    M_FRM_IFD.Tag = "Masters"
    M_FRM_IFD.Show 1
ElseIf i = 3 Then
    If MsgBox("Confirm DAS DataBase Restore from Latest AutoBackup ?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
        M_FRM_RST_BKP.Show 1
    End If
ElseIf i = 4 Then
    M_FRM_QUK_RST.Show 1
End If
End Sub

Private Sub M_SYS_IMP_TOT_Click()
M_FRM_MAS.Tag = "Complete"
M_FRM_MAS.Show 1
End Sub
Private Sub M_SYS_INI_Click()
SYS_PSW.Caption = "DAS Version 4.0"
SYS_PSW.Show 1
End Sub
Private Sub M_SYS_INV_Click(i As Integer)
If i = 0 Then
    SYS_PSW.Tag = "Template"
    SYS_PSW.Show 1
ElseIf i = 1 Then
    USR_FRM.Show 1
ElseIf i = 2 Then
    SYS_PSW.Tag = "Module"
    SYS_PSW.Show 1
ElseIf i = 3 Then
    SHT_RSN.Show 1
End If
End Sub
Private Sub M_SYS_RPC_Click()
Shell App.Path & "\Rapidrun.exe"
'M_FRM_RPT.Show 1
End Sub
Private Sub M_RPT_COR_PUR_MAR_RPT_Click(i As Integer)
If i = 0 Then
    OpenReport CURDIR & "\REP_LIB.RP1", "Margin Report", 0, "", 6
ElseIf i = 1 Then
    M_PUR_PQY.Tag = "PUR_SUM_VWT"
    M_PUR_PQY.Show 1
End If
End Sub

Private Sub M_SYS_UPD_STR_Click()
On Error Resume Next
con.Execute "ALTER TABLE VAN ADD COLUMN TRP_IDY N(3) NULL"
con.Execute "ALTER TABLE TRT ALTER COLUMN T_DAY_STK NULL"
con.Execute "ALTER TABLE TRT ALTER COLUMN T_WEK_STK NULL"
con.Execute "ALTER TABLE TRT ALTER COLUMN T_MTH_STK NULL"
con.Execute "ALTER TABLE TRT ALTER COLUMN A_DAY_STK NULL"
con.Execute "ALTER TABLE TRT ALTER COLUMN A_WEK_STK NULL"
con.Execute "ALTER TABLE TRT ALTER COLUMN A_MTH_STK NULL"
con.Execute "ALTER TABLE TRT ALTER COLUMN MTH_DYS NULL"
con.Execute "ALTER TABLE TRT ALTER COLUMN WEK_DYS NULL"
con.Execute "ALTER TABLE TRT ALTER COLUMN CMP_WEK NULL"
con.Execute "ALTER TABLE TRT ALTER COLUMN CMP_MTH NULL"
End Sub
Private Sub M_SYS_UPW_Click()
USR_FRM.Show 1
End Sub
Private Sub MDIForm_Activate()
On Error GoTo lab
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select dir_msg from dir where dir_idy like 'BMP' AND DIR_IDX LIKE 'IDY'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    If Dir(Trim(rs("DIR_MSG"))) = "" Then
        con.Execute "UPDATE DIR WHERE DIR_IDY LIKE 'BMP' AND DIR_IDX LIKE 'IDY' SET DIR_MSG = ''"
    Else
        Picture1.Picture = LoadPicture(Trim(rs("DIR_MSG")))
        If Len(Trim(rs!DIR_MSG)) <> 0 Then M_SYS_BMP(0).Checked = True
    End If
End If
Exit Sub
lab:
MsgBox "Bitmap Display Error"
End Sub

Private Sub MDIForm_Resize()
Picture1.BorderStyle = 0
'Picture1.BackColor = vbRed
If (Me.Height - STS_BAR.Height - 750) > 0 Then Picture1.Height = Me.Height - STS_BAR.Height - 750
If (Me.Width - Label1.Width - 200) > 0 Then Label1.Left = Me.Width - Label1.Width - 200
If (Label1.Left + ((Label1.Width - Image1.Width) / 2)) > 0 Then Image1.Left = Label1.Left + ((Label1.Width - Image1.Width) / 2)
If (Me.Height - STS_BAR.Height - 1550) > 0 Then Label1.Top = Me.Height - STS_BAR.Height - 1550
If (Label1.Top - Label1.Height - Image1.Height - 50) > 0 Then Image1.Top = Label1.Top - Label1.Height - Image1.Height - 50
'Picture1.Width = Me.Width
End Sub

Private Sub MDIForm_QueryUnload(cancel As Integer, UnloadMode As Integer)
LOGOUT
End Sub
Private Sub MDIForm_Unload(cancel As Integer)
LOGOUT
End Sub
Private Sub mnu_bill_Click()
M_FRM_IRT.Show 1
End Sub
Private Sub MDIForm_Load()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from fun", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    F2 = rs(0) & ""
    F3 = rs("FUN_THR") & ""
    F4 = rs("FUN_FOR") & ""
    F5 = rs("fun_fiv") & ""
    F6 = rs("fun_six") & ""
    F7 = rs("fun_sev") & ""
    F8 = rs("fun_eig") & ""
End If
Load M_FIL_SCH(1)
Load M_FIL_SCH(2)
Load M_FIL_SCH(3)
Load M_FIL_SCH(4)
Load M_FIL_SCH(5)
Load M_FIL_SCH(6)
Load M_FIL_SCH(7)
Load M_FIL_SCH(8)
Load M_FIL_SCH(9)
Load M_FIL_SCH(10)
'Load M_FIL_SCH(11)
'Load M_FIL_SCH(12)
'Load M_FIL_SCH(13)
'Load M_FIL_SCH(14)
'Load M_FIL_SCH(15)
'Load M_FIL_SCH(16)
M_FIL_SCH(1).Caption = "&2 Schemes  - HighVolume and Key Accounts"
M_FIL_SCH(2).Caption = "&3 Expenses - Claim TopSheets and Submission"
M_FIL_SCH(2).Visible = False
M_FIL_SCH(2).Enabled = False
M_FIL_SCH(3).Caption = "-"
M_FIL_SCH(4).Caption = "&4 SKU Rates - Institutions and Key Accounts"
M_FIL_SCH(5).Caption = "-"
M_FIL_SCH(6).Caption = "&5 Sales Calendar - Closure, Cycle, Reports"
M_FIL_SCH(7).Caption = "-"
M_FIL_SCH(8).Caption = "&6 Route Salesmen - Commissions and Incentives"
M_FIL_SCH(9).Caption = "&7 Route Salesmen - Targets vs Performance"
M_FIL_SCH(10).Caption = "&8 Route Salesman - Beat Plan, Route Sequence"
'M_FIL_SCH(11).Caption = "-"
'M_FIL_SCH(12).Caption = "&9 GSM Modem SMS Send Interface"
'M_FIL_SCH(13).Caption = "-"
'M_FIL_SCH(14).Caption = "&a Accounts and Audit Interface"
'M_FIL_SCH(15).Caption = "&b BarCode LabelPrint Interface"
'M_FIL_SCH(16).Caption = "&c Retail Outlet GIS Lat, Long Coordinates"
Load OLXRF(1)
Load OLXRF(2)
'Load OLXRF(3)
Load OLXRF(4)
Load OLXRF(5)
'Load OLXRF(6)
'Load OLXRF(7)
'Load OLXRF(8)
'Load OLXRF(9)
OLXRF(1).Caption = "&2 Purchase Order Entry"
OLXRF(2).Caption = "-"
OLXRF(4).Caption = "&4 Purchase Order Verification"
OLXRF(5).Caption = "&5 Supplier Order Sheet"
'OLXRF(6).Caption = "-"
'OLXRF(7).Caption = "&6 Sales Orders Xrf Invoices"
'OLXRF(8).Caption = "&7 Load Sheet Xrf Van Challan"
Load M_MDI.TCO_ENT(1)
Load M_COR_FRE_SAL(1)
M_COR_FRE_SAL(1).Caption = "-"
Load M_COR_VAN_VVR(1)
M_COR_VAN_VVR(1).Caption = "-"
Load DIS_STK_REP(1)
DIS_STK_REP(1).Caption = "&k Net Load Stock Movement"
Load DIS_STK_REP(2)
DIS_STK_REP(2).Caption = "-"
Load DIS_STK_REP(3)
DIS_STK_REP(3).Caption = "&l SKU - Stock Breakup Report"
Load DIS_STK_REP(4)
DIS_STK_REP(4).Caption = "&m Non Salable Product Details"
Load DIS_STK_REP(5)
DIS_STK_REP(5).Caption = "&n Non Salable Product Details - VAN"
Load M_RPT_OTH_DIP(1)
M_RPT_OTH_DIP(1).Caption = "&b Product Discount Retailer Wise"
Load M_RPT_OTH_DIP(2)
M_RPT_OTH_DIP(2).Caption = "&c Product Discount Segment Wise"
Load M_RPT_OTH_DIP(3)
M_RPT_OTH_DIP(3).Caption = "&d Product Discount Supplier Wise"
'Load M_COR_STK_STO(1)          'CHANGED BY PNR AS PER PRINCE CHOPRA EMAIL
'M_COR_STK_STO(1).Caption = "-" 'CHANGED BY PNR AS PER PRINCE CHOPRA EMAIL
Load M_RPT_OTH_CLM(2)
M_RPT_OTH_CLM(2).Caption = "&f Schemes Report - Brand Volume"
Load M_RPT_OTH_CLM(3)
M_RPT_OTH_CLM(3).Caption = "&g Schemes Report - Segment Value"
Load M_RPT_OTH_CLM(1)
M_RPT_OTH_CLM(1).Caption = "&e Schemes Report - SKU"
Load M_SYS_IMP_MAS(1)
M_SYS_IMP_MAS(1).Caption = "&Export-Import File Data"
Load M_SYS_IMP_MAS(2)
M_SYS_IMP_MAS(2).Caption = "&Import DAS DataBase"
Load M_SYS_IMP_MAS(3)
M_SYS_IMP_MAS(3).Caption = "&Restore DAS AutoBackup - Corrupted DB"
Load M_SYS_IMP_MAS(4)
M_SYS_IMP_MAS(4).Caption = "&QuickDB DAS Restore - Power Failure"
Load M_COR_SAL_LDS(1)
M_COR_SAL_LDS(1).Caption = "-"
Load M_RPT_ACT_FAC_GTBL(1)
M_RPT_ACT_FAC_GTBL(1).Caption = "-"
Load M_RPT_ACT_FAC_GTBL(2)
M_RPT_ACT_FAC_GTBL(2).Caption = "ROI Report"
Load M_MIR_PPV(1)
M_MIR_PPV(1).Caption = "-"
Load M_MIR_ROU(1)
M_MIR_ROU(1).Caption = "-"
Load M_RPT_COR_STK_ANL_MVR(1)
M_RPT_COR_STK_ANL_MVR(1).Caption = "-"
Load M_SYS_INV(1)
M_SYS_INV(1).Caption = "&Users and Passwords"
Load M_SYS_INV(2)
M_SYS_INV(2).Caption = "&Module Check"
M_SYS_INV(2).Visible = False
Load M_SYS_INV(3)
M_SYS_INV(3).Caption = "ShutD&ownStatus"
Load M_SYS_BMP(1)
Load M_SYS_ACP(1)
M_SYS_ACP(1).Caption = "&Report Sets"
M_SYS_BMP(1).Caption = "&Function Key Defaults"
Load M_SYS_BMP(2)
M_SYS_BMP(2).Caption = "&1 DAS Mobile - Identification"
M_SYS_BMP(2).Visible = False
Load M_SYS_BMP(3)
M_SYS_BMP(3).Caption = "&2 DAS Mobile - Retailers Information"
M_SYS_BMP(3).Visible = False
Load M_SYS_BMP(4)
M_SYS_BMP(4).Caption = "&3 DAS Mobile - Products Information"
M_SYS_BMP(4).Visible = False
Load M_SYS_BMP(5)
M_SYS_BMP(5).Caption = "&4 DAS Mobile - Data Exchange"
M_SYS_BMP(5).Visible = False
Load M_EXT_WDP(1)
M_EXT_WDP(1).Caption = "&Shutdown"
' Two Way Pack Reports
Load M_COR_TWO_REP_ESH(1)
M_COR_TWO_REP_ESH(1).Caption = "&b Returnables ROL COL Variance by Pack-Shell"
Load M_COR_TWO_REP_ESH(2)
M_COR_TWO_REP_ESH(2).Caption = "&c Returnables ROL COL Variance by Pack-Party"
Load M_COR_TWO_REP_ESH(3)
M_COR_TWO_REP_ESH(3).Caption = "&d Returnables ROL COL Summary by Party"
Load M_COR_TWO_REP_ESH(4)
M_COR_TWO_REP_ESH(4).Caption = "&e Returnables Float (Deposits) Account"
Load M_COR_TWO_REP_ESH(5)
M_COR_TWO_REP_ESH(5).Caption = "&f Returnables Register"
Load M_COR_TWO_REP_ESH(6)
M_COR_TWO_REP_ESH(6).Caption = "&g Returnables Stk-On-Hand Net Load"
' Adding Report Sales
Load M_RPT_COR_SAL_RPT(1)
M_RPT_COR_SAL_RPT(1).Caption = "&b Report - Retailer"
Load M_RPT_COR_SAL_RPT(2)
M_RPT_COR_SAL_RPT(2).Caption = "&c Summary"
Load M_RPT_COR_SAL_RPT(3)
M_RPT_COR_SAL_RPT(3).Caption = "&d Register - 132 Col."
Load M_RPT_COR_SAL_RPT(4)
M_RPT_COR_SAL_RPT(4).Caption = "&e Product"
Load M_RPT_COR_SAL_RPT(5)
M_RPT_COR_SAL_RPT(5).Caption = "&f Returns"
Load M_RPT_COR_SAL_RPT(6)
M_RPT_COR_SAL_RPT(6).Caption = "&g Returns Product"
Load M_RPT_COR_SAL_RPT(7)
M_RPT_COR_SAL_RPT(7).Caption = "&h Stock Type"
Load M_RPT_COR_SAL_RPT(8)
M_RPT_COR_SAL_RPT(8).Caption = "&i Stock Type Product"
Load M_RPT_COR_SAL_RPT(9)
M_RPT_COR_SAL_RPT(9).Caption = "&j Invoice Returns"
Load M_RPT_COR_SAL_RPT(10)
M_RPT_COR_SAL_RPT(10).Caption = "&k Scheme Quantity - Scheme"
Load M_RPT_COR_SAL_RPT(11)
M_RPT_COR_SAL_RPT(11).Caption = "&l Scheme Quantity - Date"
Load M_RPT_COR_SAL_RPT(12)
M_RPT_COR_SAL_RPT(12).Caption = "&m Scheme Value"
Load M_RPT_COR_SAL_RPT(13)
M_RPT_COR_SAL_RPT(13).Caption = "&n Tax Report"
Load M_RPT_COR_SAL_RPT(14)
M_RPT_COR_SAL_RPT(14).Caption = "&o Tax Product"
Load M_RPT_COR_SAL_RPT(15)
M_RPT_COR_SAL_RPT(15).Caption = "&p Total Tax Product"
Load M_RPT_COR_SAL_RPT(16)
M_RPT_COR_SAL_RPT(16).Caption = "&q SKU Lines Count - 2"
Load M_RPT_COR_SAL_RPT(17)
M_RPT_COR_SAL_RPT(17).Caption = "&r VAT Input"
Load M_RPT_COR_SAL_RPT(18)
M_RPT_COR_SAL_RPT(18).Caption = "&s VAT Output"
Load M_RPT_COR_SAL_RPT(19)
M_RPT_COR_SAL_RPT(19).Caption = "&t VAT Input-Output"
Load M_RPT_COR_SAL_RPT(20)
M_RPT_COR_SAL_RPT(20).Caption = "&u SKU UOM Summary"
Load M_RPT_COR_SAL_RPT(21)
M_RPT_COR_SAL_RPT(21).Caption = "&v Claim Detail Summary - 4"
Load M_RPT_COR_SAL_RPT(22)
M_RPT_COR_SAL_RPT(22).Caption = "&w Summary Retail Channel"
Load M_RPT_COR_SAL_RPT(23)
M_RPT_COR_SAL_RPT(23).Caption = "&x Summary Retail Location"
Load M_RPT_COR_SAL_RPT(24)
M_RPT_COR_SAL_RPT(24).Caption = "&y Summary Retail Market"
Load M_RPT_COR_SAL_RPT(25)
M_RPT_COR_SAL_RPT(25).Caption = "&z Summary Retail Route"

Load M_RPT_COR_SAL_RPT(26)
M_RPT_COR_SAL_RPT(26).Caption = "&1 Summary Retail Salesman"
Load M_RPT_COR_SAL_RPT(27)
M_RPT_COR_SAL_RPT(27).Caption = "&2 Register - 80 Column"
Load M_RPT_COR_SAL_RPT(28)
M_RPT_COR_SAL_RPT(28).Caption = "&3 Retail Discount"
Load M_RPT_COR_SAL_RPT(29)
M_RPT_COR_SAL_RPT(29).Caption = "&4 Retailer - SKU Summary"
Load M_RPT_COR_SAL_RPT(30)
M_RPT_COR_SAL_RPT(30).Caption = "&5 Brand - SKU Summary"
Load M_RPT_COR_SAL_RPT(31)
M_RPT_COR_SAL_RPT(31).Caption = "&6 Pack Size - SKU Summary"
Load M_RPT_COR_SAL_RPT(32)
M_RPT_COR_SAL_RPT(32).Caption = "&7 Pack Type - SKU Summary"
Load M_RPT_COR_SAL_RPT(33)
M_RPT_COR_SAL_RPT(33).Caption = "&8 Sales Register SoftCopy"
Load M_RPT_COR_SAL_RPT(34)
M_RPT_COR_SAL_RPT(34).Caption = "&9 Van Register SoftCopy"

Load M_RPT_COR_PUR_MAR_RPT(1)
M_RPT_COR_PUR_MAR_RPT(1).Caption = "&h Summary - Vol-Wgt"
' Stock Additions
Load M_RPT_COR_STK_ANL_AGD(1)
M_RPT_COR_STK_ANL_AGD(1).Caption = "&b SKU - Aging - Purchase Date - Cases"
Load M_RPT_COR_STK_ANL_AGD(2)
M_RPT_COR_STK_ANL_AGD(2).Caption = "&c SKU - Aging - Purchase Date - Value"
Load M_RPT_COR_STK_ANL_AGD(3)
M_RPT_COR_STK_ANL_AGD(3).Caption = "&d SKU - Aging - Mfg Date - Cases"
Load M_RPT_COR_STK_ANL_AGD(4)
M_RPT_COR_STK_ANL_AGD(4).Caption = "&e SKU - Batch Stock Summary - UPC"
Load M_RPT_COR_STK_ANL_AGD(5)
M_RPT_COR_STK_ANL_AGD(5).Caption = "&f SKU - Batch Stock Summary - Cases"
Load M_RPT_COR_STK_ANL_AGD(6)
M_RPT_COR_STK_ANL_AGD(6).Caption = "&g SKU - Summary - Cases"
Load M_RPT_COR_STK_ANL_AGD(7)
M_RPT_COR_STK_ANL_AGD(7).Caption = "-"
Load M_RPT_COR_STK_ANL_AGD(8)
M_RPT_COR_STK_ANL_AGD(8).Caption = "&h SKU - Opening Stock - by SKU"
Load M_RPT_COR_STK_ANL_AGD(9)
M_RPT_COR_STK_ANL_AGD(9).Caption = "&i Stock Inward, Receipts - Summary"
Load M_RPT_COR_STK_ANL_AGD(10)
M_RPT_COR_STK_ANL_AGD(10).Caption = "&j SKU - Stock Inward - by SKU"
Load M_RPT_COR_STK_ANL_AGD(11)
M_RPT_COR_STK_ANL_AGD(11).Caption = "&k Stock Outward, Issues - Summary"
Load M_RPT_COR_STK_ANL_AGD(12)
M_RPT_COR_STK_ANL_AGD(12).Caption = "&l SKU - Stock Outward - by SKU"
Load M_RPT_COR_STK_RGL_STL(1)
M_RPT_COR_STK_RGL_STL(1).Caption = "&n SKU Stock Ledger - Date Consolidated"
Load M_RPT_COR_STK_RGL_STL(2)
M_RPT_COR_STK_RGL_STL(2).Caption = "&o Stock Damage - Breakup - by SKU"
Load M_RPT_COR_STK_RGL_STL(3)
M_RPT_COR_STK_RGL_STL(3).Caption = "&p Stock Damage - Breakup - by SKU,Segment"
Load M_RPT_COR_STK_RGL_STL(4)
M_RPT_COR_STK_RGL_STL(4).Caption = "&q Inventory Availability Analysis"
' Others - KO Reports Addition - PNR
Load M_MIR_MBR(1)
M_MIR_MBR(1).Caption = "&b Non Buying Retailers - Past 14 Days"
Load M_MIR_MBR(2)
M_MIR_MBR(2).Caption = "&c Salesman Route Productivity - WeekDay"
Load M_MIR_MBR(3)
M_MIR_MBR(3).Caption = "&d Distributor Stock Report - Qty <= 6 Weeks to BBD"
Load M_MIR_MBR(4)
M_MIR_MBR(4).Caption = "&e Secondary Sale Summary - YTD : MTD : WTD"
Load M_MIR_MBR(5)
M_MIR_MBR(5).Caption = "&f Retailer A/c Receivable and Glass COL"
Load M_MIR_MBR(6)
M_MIR_MBR(6).Caption = "-"
Load M_MIR_MBR(7)
M_MIR_MBR(7).Caption = "&g Monthly Business Record - Figs in '000"
Load M_MIR_MBR(8)
M_MIR_MBR(8).Caption = "&h Salesman Turnover and Collection Report "
Load M_MIR_MBR(9)
M_MIR_MBR(9).Caption = "&i Visit Counts - Invoice Numbers-Market"
Load M_MIR_MBR(10)
M_MIR_MBR(10).Caption = "-"
Load M_MIR_MBR(11)
M_MIR_MBR(11).Caption = "&j Scheme-Claim Qty Report-MER Sorted - 4 Rpts"
Load M_MIR_MBR(12)
M_MIR_MBR(12).Caption = "&k Scheme-Claim Value Report-MER Sorted"
Load M_MIR_MBR(13)
M_MIR_MBR(13).Caption = "&l HVO OnSpot Report"
Load M_MIR_MBR(14)
M_MIR_MBR(14).Caption = "&m HVO Periodic Report"
Load M_MIR_MBR(15)
M_MIR_MBR(15).Caption = "-"
Load M_MIR_MBR(16)
M_MIR_MBR(16).Caption = "&n Empties Movement Report"
M_MIR_MBR(16).Enabled = False
Load M_MIR_MBR(17)
M_MIR_MBR(17).Caption = "&o Retailer Net COL-ROL, Net AR-AP Report"
M_MIR_MBR(17).Enabled = False
Load M_MIR_MBR(18)
M_MIR_MBR(18).Caption = "&p Salesman Commission Report-Van Challan"
M_MIR_MBR(18).Enabled = False
Load M_MIR_MBR(19)
M_MIR_MBR(19).Caption = "&q Route Gross Profitability Report"
M_MIR_MBR(19).Enabled = False
Load M_MIR_MBR(20)
M_MIR_MBR(20).Caption = "&r Godown Gross Profitability Report"
M_MIR_MBR(20).Enabled = False
Load M_MIR_MBR(21)
M_MIR_MBR(21).Caption = "-"
Load M_MIR_MBR(22)
M_MIR_MBR(22).Caption = "&s Fill Rate Summary Report"
'M_MIR_MBR(22).Enabled = False
Load M_MIR_MBR(23)
M_MIR_MBR(23).Caption = "&t Fill Rate Detailed Report"
'M_MIR_MBR(23).Enabled = False

' KO Reports Addition End - PNR

'My Page Reports Addition - PNR
Load M_RPT_MYPAGE_SQA(1)
M_RPT_MYPAGE_SQA(1).Caption = "&b Sales Value Analysis ReadyStock"
Load M_RPT_MYPAGE_SQA(2)
M_RPT_MYPAGE_SQA(2).Caption = "&c Funds Analysis ReadyStock"
Load M_RPT_MYPAGE_SQA(3)
M_RPT_MYPAGE_SQA(3).Caption = "&d Vehicle Utlization by SalesMan"
Load M_RPT_MYPAGE_SQA(4)
M_RPT_MYPAGE_SQA(4).Caption = "&e Vehicle Utlization by Vehicle"
Load M_RPT_MYPAGE_SQA(5)
M_RPT_MYPAGE_SQA(5).Caption = "&f Status Sheet"
'My Page Reports Addition End - PNR

M_FIL_CAT.Tag = "Categories"
M_FIL_LOC.Tag = "Locations"
M_FIL_SMN.Tag = "Salesman"
M_FIL_MAR.Tag = "Markets"
M_FIL_GRP.Tag = "Group Accounts"
M_FIL_ACT.Tag = "Accounts & Parties"
M_FIL_PGR.Tag = "Product Categories"
M_FIL_PRD.Tag = "Products (SKU)"
'M_FIL_RAT.Tag = "Key Accounts SKU Rates"
'M_FIL_SCH.Tag = "Schemes - Trade"
'M_FIL_CLD.Tag = "Sales Calender"
'M_FIL_TGT.Tag = "Targets - Salesmen"
'M_FIL_RUT.Tag = "Route Sequence"
'M_FIL_SIC(0).Tag = "Commissions"
'M_FIL_HYP5(0).Tag = "Day ClosUre n Data XChange"
'M_FIL_DAS_CRS(0).Tag = "DAS-CRS"
M_COR.Tag = "Core"
M_COR_PUR_SSL.Tag = "Purchases"
M_COR_SAL.Tag = "Sales"
M_COR_FRE_SAL(0).Tag = "Credit Sales"
M_COR_SAL_RSB(0).Tag = "Ready Stock Bills"
M_COR_VAN_VVR(0).Tag = "Van Variance Check"
mnu_bill.Tag = "Invoices vs Receipts"
M_COR_SAL_LDS(0).Tag = "Load Sheet"
M_COR_PSH_SAL.Tag = "Push Cart Sales"
M_COR_STK.Tag = "Stocks"
M_COR_STK_OPS.Tag = "Opening Stock"
M_COR_STK_CRN.Tag = "Credit Notes"
M_COR_STK_DRN.Tag = "Debit Notes"
M_COR_STK_STI.Tag = "Transfer Inward"
M_COR_STK_STO(0).Tag = "Transfer Outward"
M_COR_STK_ADJ_JRN.Tag = "Stock Adjustment"
M_COR_STK_COV.Tag = "Stock Conversion"
'mnu_cmp.Tag = "Supplier Order"
M_MNU_SHL_SNP1.Tag = "snp"
M_COR_HYP5(0).Tag = "twp"
TCO_ENT(0).Tag = "tent"
TCO_ENT(1).Tag = "ti"
M_MNU_SHL_SNP.Tag = "Shells n Packs"
M_MNU_SHL_INW(0).Tag = "Shells Additionals"
M_COR_SHL_TRS.Tag = "Shell Transactions"
M_COR_UPS.Tag = "Update Stocks"
M_ACT.Tag = "Accounts"
M_ACT_RCT.Tag = "Receipts"
M_ACT_PMT.Tag = "Payments"
M_ACT_JRN.Tag = "Journals"
M_ACT_VCH_JVN.Tag = "Journal Entries"
M_ACT_VCH_CRN.Tag = "Credit Note"
M_ACT_VCH_DRN.Tag = "Debit Note"
M_ACT_VCH_OPB.Tag = "Opening Balance"
M_AST.Tag = "Assets"
M_FIL_ATT.Tag = "Asset Type"
M_FIL_ASM.Tag = "Asset Details"
M_AST_INW.Tag = "Inward Entry"
M_AST_OUT.Tag = "Outward Entry"
m_mir.Tag = "Analysis"
ANA_RET.Tag = "Retailers"
M_MIR_TOT.Tag = "Ranking"
M_ANA_RET_LBD.Tag = "Last Billing Date"
M_ANA_RET_RRN.Tag = "Retailers Route NB"
M_ANA_RET_RSN.Tag = "Retailers Salesman NB"
M_MIR_NPD.Tag = "Non Operational"
M_MIR_RCG.Tag = "Categories Count"
M_MIR_SOT.Tag = "Sales Pareto"
M_MIR_CFR.Tag = "Coverage Frequency"
M_MIR_YSR.Tag = "Yearly Sales"
ANA_SKU.Tag = "Products(SKU)"
M_MIR_PPD.Tag = "SKU Pareto"
M_MIR_PPD_SEG.Tag = "SKU Segment Pareto"
M_MIR_PPD_CAT.Tag = "SKU Category Pareto"
M_MIR_PPV(0).Tag = "Sale Value Pareto"
M_MIR_SPD.Tag = "Retailers Pareto"
m_mir_cat.Tag = "Categories Pareto"
m_mar_loc.Tag = "Locations Pareto"
M_MIR_MAR.Tag = "Markets Pareto"
M_MIR_ROU(0).Tag = "Routes Pareto"
M_MIR_RAR.Tag = "Penetration Count"
mnu_sup.Tag = "Sales Report - Splr"
mnu_stk.Tag = "Sales Report - Stock"
M_RPT_COR_STK_ANL_MVR(0).Tag = "Product Sales"
M_MIR_YQP.Tag = "Yearly Sales - Units"
M_MIR_YQC.Tag = "Yearly Sales - Cases"
mnu_sup_b.Tag = "Actual Sales Report"
ANA_OTH.Tag = "Others"
'''''M_MIR_MBR.Tag = "Month Wise Business Record"
'''''M_MIR_STR.Tag = "Salesman Turnover Report"
'''''M_MIR_STR_MAR_MCF.Tag = "Market Coverage Frequency"

'M_SYS.Tag = "SYSTEM"
SYS_SET.Tag = "Settings"
M_SYS_CLG.Tag = "CLI"
SYS_DTB.Tag = "Database"
M_SYS_BKP.Tag = "BKP"
SYS_MIS.Tag = "DAS Mobile and Others"
M_SYS_BMP(0).Tag = "BI"
M_SYS_BMP(1).Tag = "FKD"
M_SYS_IMP_MAS(0).Tag = "IED"
M_SYS_IMP_MAS(1).Tag = "EFD"
M_SYS_IMP_MAS(2).Tag = "VDB"
M_SYS_IMP_MAS(3).Tag = "RDB"
M_SYS_IMP_MAS(4).Tag = "QDB"
M_SYS_CCP.Tag = "CSW"
M_SYS_INI.Tag = "DF"
M_SYS_ACP(0).Tag = "AFY"
M_SYS_ACP(1).Tag = "RS"
M_SYS_INV(3).Tag = "SDS"
M_SYS_INV(1).Tag = "UAP"
M_SYS_INV(2).Tag = "MC"
M_SYS_INV(0).Tag = "SD"



M_SYS_BMP(2).Tag = "DAS Mobile - Identification"
M_SYS_BMP(3).Tag = "DAS Mobile - Retailers Information"
M_SYS_BMP(4).Tag = "DAS Mobile - Products Information"
M_SYS_BMP(5).Tag = "DAS Mobile - Data Exchange"

M_ANA_SPC_RPT.Tag = "Specific Reports"
VAN_VAR.Tag = "Van Variance"
SHL_VAR.Tag = "Shell Variance"
VAN_SHL_VAR.Tag = "Van Shell Variance"
SMN_CUM_BAL.Tag = "Smn Cumulative Balance"
DLY_PRD_SUM.Tag = "Daily Productivity Summary"
DIS_SCH_CPC.Tag = "Discount Scheme CPC"
VEH_UTL.Tag = "Vehicle Utilization"
VAN_CHL_STS.Tag = "Van Challans Status"
DIS_STK_REP(0).Tag = "Distributors Stock Report"

M_ANA_PRG_RPT.Tag = "Progress Reports"
M_ANA_PRG_CAT.Tag = "Category Sales Progress"
M_ANA_PRG_LOC.Tag = "Location Sales Progress"
M_ANA_PRG_MAR.Tag = "Market Sales Progress"
M_ANA_PRG_SMN.Tag = "Salesman Sales Progress"
M_ANA_PRG_ROU.Tag = "Route Sales Progress"
M_ANA_PRG_RET.Tag = "Retailer Sales Progress"
M_ANA_COM_RPT.Tag = "Combination Reports"
End Sub
Private Sub mnu_cmp_Click(i As Integer)
If i = 1 Then
    M_FRM_ORD.Show 1
End If
End Sub
Private Sub OLXRF_Click(i As Integer)
If i = 0 Then
    M_FRM_SOR.Show 1
ElseIf i = 1 Then
    M_FRM_POR.Show 1
ElseIf i = 3 Then
    M_FRM_LDSORD.Show 1
ElseIf i = 5 Then
    M_FRM_ORD.Show 1
End If
End Sub

Private Sub mnu_stk_Click()
M_REP_PQY.Tag = "stk"
M_REP_PQY.Show 1
End Sub

Private Sub mnu_sup_b_Click()
M_REP_PQY.Tag = "sup_b"
M_REP_PQY.Show 1
End Sub

Private Sub mnu_sup_Click()
M_REP_PQY.Tag = "sup"
M_REP_PQY.Show 1
End Sub
Private Sub REC_STM_Click()
M_FRM_DBT.Show 1
End Sub
Private Sub M_RPT_ACT_DCR_Click()
M_FRM_DBT.Tag = "DCR_STM"
M_FRM_DBT.Show 1
End Sub
Private Sub M_RPT_ACT_CIB_Click()
M_FRM_DBT.Tag = "DCR_CIB"
M_FRM_DBT.Show 1
End Sub
Private Sub M_RPT_ACT_DCR_SCR_Click()
M_FRM_DBT.Tag = "DCR_SCR"
M_FRM_DBT.Show 1
End Sub
Private Sub SHL_VAR_Click()
M_REP_SMN.Tag = "SHL"
M_REP_SMN.Show 1
End Sub
Private Sub SMN_CUM_BAL_Click()
M_REP_SMN.Tag = "CUM"
    M_REP_SMN.Show 1
End Sub
Private Sub SYS_SHT_STS_Click()
SHT_RSN.Show 1
End Sub
Private Sub TCO_ENT_Click(i As Integer)
'If i = 0 Then
'M_FRM_RTG.Show 1
'ElseIf i = 1 Then
''M_FRM_MBL.Show 1
'End If
End Sub
Private Sub VAN_CHL_STS_Click()
M_REP_VAN.Tag = "CHL_STS"
M_REP_VAN.Show 1
End Sub
Private Sub VAN_SHL_VAR_Click()
M_REP_SMN.Tag = "VSV"
M_REP_SMN.Show 1
End Sub
Private Sub VAN_VAR_Click()
    M_REP_SMN.Tag = "VAN"
    M_REP_SMN.Show 1
End Sub

Private Sub VEH_UTL_Click()
M_FRM_ASD.Tag = "MyPage_VUT"
M_FRM_ASD.Show 1
'M_REP_VAN.Tag = "VUR"      'Old Vehicle Report specific for one vehicle is removed. PNR
'M_REP_VAN.Show 1
End Sub
