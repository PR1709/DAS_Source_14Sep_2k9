VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmNewRateTemplate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Price List Update and Batch Values Edit"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7890
   Icon            =   "frmNewRateTemplate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7890
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1778
      TabIndex        =   30
      Top             =   5085
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
      Height          =   375
      Left            =   2898
      TabIndex        =   31
      Top             =   5085
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   4018
      TabIndex        =   32
      Top             =   5085
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5138
      TabIndex        =   33
      Top             =   5085
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   320
      Index           =   23
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   320
      Index           =   22
      Left            =   5880
      TabIndex        =   29
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   21
      Left            =   4920
      TabIndex        =   28
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   20
      Left            =   3960
      TabIndex        =   27
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   19
      Left            =   3000
      TabIndex        =   26
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   18
      Left            =   2040
      TabIndex        =   25
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   17
      Left            =   1080
      TabIndex        =   24
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   16
      Left            =   120
      TabIndex        =   23
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   15
      Left            =   6840
      TabIndex        =   22
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   14
      Left            =   5880
      TabIndex        =   21
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   13
      Left            =   4920
      TabIndex        =   20
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   12
      Left            =   3960
      TabIndex        =   19
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   11
      Left            =   3000
      TabIndex        =   18
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   10
      Left            =   2040
      TabIndex        =   17
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   9
      Left            =   1080
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   8
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   7
      Left            =   6840
      TabIndex        =   14
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   6
      Left            =   5880
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   5
      Left            =   4920
      TabIndex        =   12
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   320
      Index           =   4
      Left            =   3960
      TabIndex        =   11
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   320
      Index           =   3
      Left            =   3000
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   320
      Index           =   2
      Left            =   2040
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   320
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Import"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Product Short Name"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Product - SKU Name - Part Number"
      Height          =   375
      Index           =   1
      Left            =   2978
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      Caption         =   "SKU Supplier ID"
      Height          =   375
      Index           =   2
      Left            =   6120
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   320
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   3480
      Width           =   975
   End
   Begin MSComctlLib.ListView lvBatch 
      Height          =   1425
      Left            =   2880
      TabIndex        =   7
      Top             =   1605
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2514
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   25
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Part Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PurRate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SaleRate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "MRP"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "PurIncDis"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "SalIncDis"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "PurVAT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "PurCrateRental"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "PurCrateVAT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "PurTOTVAT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "PurLPP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "UPurLPP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "SaleVAT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "SaleCrateRental"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "SaleCrateVAT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "SaleTOTVAT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "SaleLPP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "USalLPP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "SKU_SBU"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "BaseCrateRate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "BaseCrateDis"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "BasePurRate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "BaseSaleTax"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Batch Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Qty"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtTemplate 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2978
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdBarCodeTemplate 
      Caption         =   "..."
      Height          =   375
      Left            =   6488
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin MSComDlg.CommonDialog dlgTemplate 
      Left            =   7320
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "New Rate Template"
      Filter          =   "*.xls"
      InitDir         =   "C:\Program Files\EFFMCG4"
   End
   Begin VB.ListBox lstprd_idy 
      Height          =   2790
      Left            =   8160
      TabIndex        =   36
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox lstprd_nme 
      Height          =   1230
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label lblProcess 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Qty"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   22
      Left            =   6840
      TabIndex        =   67
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label18 
      Height          =   255
      Left            =   1200
      TabIndex        =   65
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label17 
      Height          =   255
      Left            =   120
      TabIndex        =   64
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BaseSaleTax"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   8
      Left            =   5880
      TabIndex        =   48
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BasePurTax"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   23
      Left            =   4920
      TabIndex        =   61
      Top             =   4440
      Width           =   1020
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BaseCrateDis"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   9
      Left            =   3960
      TabIndex        =   49
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PurTOTVAT"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   11
      Left            =   1080
      TabIndex        =   63
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BaseCrateRate"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   21
      Left            =   3000
      TabIndex        =   60
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SKU_SBU"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   20
      Left            =   2040
      TabIndex        =   59
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USalLPP"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   19
      Left            =   1080
      TabIndex        =   58
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SaleLPP"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   18
      Left            =   120
      TabIndex        =   57
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SaleTOTVAT"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   17
      Left            =   6840
      TabIndex        =   56
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SaleCrateVAT"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   16
      Left            =   5880
      TabIndex        =   55
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SalCrateRent"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   15
      Left            =   4935
      TabIndex        =   54
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SaleVAT"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   14
      Left            =   3960
      TabIndex        =   53
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UPurLPP"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   13
      Left            =   3000
      TabIndex        =   52
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PurLPP"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   12
      Left            =   2040
      TabIndex        =   51
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PurCrateVAT"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   10
      Left            =   120
      TabIndex        =   50
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PurCrateRent"
      ForeColor       =   &H80000002&
      Height          =   255
      Index           =   7
      Left            =   6840
      TabIndex        =   47
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PurVAT"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   6
      Left            =   5880
      TabIndex        =   46
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SalIncDis"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   5
      Left            =   4920
      TabIndex        =   45
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PurIncDis"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   4
      Left            =   3960
      TabIndex        =   44
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MRPRate"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   3
      Left            =   3000
      TabIndex        =   43
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SaleRate"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   2
      Left            =   2040
      TabIndex        =   42
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PurRate"
      ForeColor       =   &H80000002&
      Height          =   225
      Index           =   1
      Left            =   1080
      TabIndex        =   41
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblPartName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3960
      TabIndex        =   40
      Top             =   1200
      Width           =   2565
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Batches for"
      Height          =   195
      Left            =   2880
      TabIndex        =   39
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Rate .XLS File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1095
      TabIndex        =   38
      Top             =   210
      Width           =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   7920
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SKU"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   3240
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&SKU Name"
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   1200
      Width           =   795
   End
End
Attribute VB_Name = "frmNewRateTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim foxprofile As String
Dim rec As New ADODB.Recordset
'Dim con As Connection


Private Sub Form_Activate()
    DoEvents
    Set rec = New ADODB.Recordset
    rec.Open "select prd_nme,prd_idy from prd WHERE CUR_LEV > 0 ORDER BY PRD_IDY", con, adOpenKeyset, adLockPessimistic
    lstprd_nme.CLEAR
    lstprd_idy.CLEAR
    While Not rec.EOF
        lstprd_nme.AddItem rec("prd_nme")
        lstprd_idy.AddItem rec("prd_idy")
        rec.MoveNext
    Wend
    If lstprd_nme.ListCount > 0 Then
        lstprd_nme.ListIndex = -1
    End If
'   rec.Close
     lblProcess.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
If KeyCode = vbKeyE And Shift = vbAltMask Then
    cmdBarCodeTemplate_Click
End If
End Sub

Private Sub Form_Load()
'    foxprofile = "C:\MAHAVIR0809"
'    Set con = New Connection
'    con = "Driver=Microsoft Visual FoxPro Driver;SourceType=DBf;SourceDB=" & foxprofile & ";BackgroundFetch=No;"
'    con.Open
'

End Sub

Private Sub Timer1_Timer()
If (lblProcess.Visible) Then
    lblProcess.Visible = False
Else
    lblProcess.Visible = True
    
End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)
SendKeys "{Home}+{End}"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        lstprd_nme.SetFocus
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub txtSearch_Change()
    Dim status As Boolean
    status = False
    For i = 0 To lstprd_nme.ListCount - 1
        If Mid(RTrim(lstprd_nme.List(i)), 1, Len(txtSearch.Text)) = UCase(txtSearch.Text) Then
            lstprd_nme.ListIndex = i
            status = True
            Exit For
        End If
    Next
    If Not status Then
        lstprd_nme.ListIndex = -1
        lstprd_idy.ListIndex = -1
    End If
End Sub

Private Sub lstprd_nme_Click()
Dim Item As ListItem
    If lstprd_nme.ListIndex < 0 Then Exit Sub
    lvBatch.ListItems.CLEAR
    lstprd_idy.ListIndex = lstprd_nme.ListIndex
    lblPartName.Caption = lstprd_nme.Text
    If rec.State = 1 Then rec.Close
    rec.Open "SELECT * from brd , prd where brd.prd_idy = prd.prd_idy and prd.prd_idy = '" & Trim(lstprd_idy.Text) & "' AND BRD.PRD_QTY > 0", con, adOpenKeyset, adLockPessimistic
    While Not rec.EOF
        Set Item = lvBatch.ListItems.Add(, , rec("prd_prn"))
        Item.ListSubItems.Add , , DecimalValue(rec("prd_pdr"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("prd_sdr"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("prd_mrp"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("p_inc_dis"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("s_inc_dis"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("p_pur_vat"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("p_crt_rnt"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("p_crt_vat"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("p_tot_vat"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("pur_lpp"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("u_pur_lpp"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("s_sal_vat"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("s_crt_rnt"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("s_crt_vat"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("s_tot_vat"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("sal_lpp"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("u_sal_lpp"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("sku_sbu"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("b_crt_rat"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("b_crt_dis"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("b_pur_tax"), 2)
        Item.ListSubItems.Add , , DecimalValue(rec("b_sal_tax"), 2)
        Item.ListSubItems.Add , , rec("bat_idy")
        If IsNull(rec("prd_qty")) = True Then
            Item.ListSubItems.Add , , "0.000"
        Else
            Item.ListSubItems.Add , , DecimalValue(rec("prd_qty"), 2)
        End If
        rec.MoveNext
    Wend
    rec.Close
    For i = 0 To Text1.COUNT - 1
        Text1(i).Text = ""
    Next
    If lvBatch.ListItems.COUNT > 0 Then
        lvBatch_ItemClick lvBatch.ListItems(1)
    End If
End Sub

Private Sub lstprd_nme_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lvBatch.SetFocus
End Sub

Private Sub lvBatch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1(0).Text = Item
    Text1(1).Text = Format(Item.ListSubItems(1), "#########0.00")
    Text1(2).Text = Format(Item.ListSubItems(2), "#########0.00")
    Text1(3).Text = Format(Item.ListSubItems(3), "#########0.00")
    Text1(4).Text = Format(Item.ListSubItems(4), "#########0.00")
    Text1(5).Text = Format(Item.ListSubItems(5), "#########0.00")
    Text1(6).Text = Format(Item.ListSubItems(6), "#########0.00")
    Text1(7).Text = Format(Item.ListSubItems(7), "#########0.00")
    Text1(8).Text = Format(Item.ListSubItems(8), "#########0.00")
    Text1(9).Text = Format(Item.ListSubItems(9), "#########0.00")
    Text1(10).Text = Format(Item.ListSubItems(10), "#########0.00")
    Text1(11).Text = Format(Item.ListSubItems(11), "#########0.00")
    Text1(12).Text = Format(Item.ListSubItems(12), "#########0.00")
    Text1(13).Text = Format(Item.ListSubItems(13), "#########0.00")
    Text1(14).Text = Format(Item.ListSubItems(14), "#########0.00")
    Text1(15).Text = Format(Item.ListSubItems(15), "#########0.00")
    Text1(16).Text = Format(Item.ListSubItems(16), "#########0.00")
    Text1(17).Text = Format(Item.ListSubItems(17), "#########0.00")
    Text1(18).Text = Format(Item.ListSubItems(18), "#########0.00")
    Text1(19).Text = Format(Item.ListSubItems(19), "#########0.00")
    Text1(20).Text = Format(Item.ListSubItems(20), "#########0.00")
    Text1(21).Text = Format(Item.ListSubItems(21), "#########0.00")
    Text1(22).Text = Format(Item.ListSubItems(22), "#########0.00")
    Text1(23).Text = Format(Item.ListSubItems(24), "#########0.000")
End Sub

Private Sub lvBatch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lvBatch_ItemClick lvBatch.SelectedItem
    Text1(1).SetFocus
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index) = DecimalValue(Text1(Index).Text, 2)
End Sub

Private Sub TEXT1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Text1(Index).Text = "" Then
        Text1(Index).Text = "0"
        End If
       Text1(Index).Text = Format(Text1(Index).Text, "##########0.00")
       Index = Index + 1
       If Text1.COUNT > Index Then
       Text1(Index).SetFocus
       End If
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = NumberFormat(KeyAscii, Text1(Index).Text)
End Sub

Private Sub cmdBarCodeTemplate_Click()
  On Error GoTo ErrHandler
  dlgTemplate.CancelError = True
  dlgTemplate.Flags = cdlOFNHideReadOnly
  dlgTemplate.filter = "New Rate Template (*.xls)|*.xls"
  dlgTemplate.FilterIndex = 2
  dlgTemplate.ShowOpen
  If dlgTemplate.FileName <> "" Then
    txtTemplate.Text = dlgTemplate.FileName
  End If
ErrHandler:
End Sub
Private Sub Command1_Click()
    Dim BatchID As String
    Dim sqlstatement As String
    If (lstprd_nme.Text) = "" Then
        MsgBox "          Product - SKU Not Selected            ", vbExclamation, "DAS Version 4.0"
        lstprd_nme.SetFocus
        Exit Sub
    End If
    If MsgBox("Create New Batch ID ?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbNo Then Exit Sub
    BatchID = "BAT" & T7increment("idy_bat")
    sqlstatement = "insert into brd (prd_qty, prd_idy , bat_idy, bat_nme,  prd_pdr, prd_sdr, prd_mrp, p_inc_dis, s_inc_dis, p_pur_vat, p_crt_rnt, p_crt_vat, p_tot_vat, pur_lpp, u_pur_lpp, s_sal_vat, s_crt_rnt, s_crt_vat, s_tot_vat, sal_lpp, u_sal_lpp, sku_sbu, b_crt_rat, b_crt_dis, b_pur_tax, b_sal_tax) values ("
    sqlstatement = sqlstatement & "0.00,   '" & lstprd_idy.Text & "','" & BatchID & "','" & "MRP : " & CStr(Text1(3).Text) & "'," & _
    DecimalValue(Text1(1).Text, 2) & "," & DecimalValue(Text1(2).Text, 2) & "," & DecimalValue(Text1(3).Text, 2) & "," & DecimalValue(Text1(4).Text, 2) & "," & DecimalValue(Text1(5).Text, 2) & "," & DecimalValue(Text1(6).Text, 2) & "," & DecimalValue(Text1(7).Text, 2) & "," & DecimalValue(Text1(8).Text, 2) & "," & DecimalValue(Text1(9).Text, 2) & "," & DecimalValue(Text1(10).Text, 2) & "," & DecimalValue(Text1(11).Text, 2) & "," & DecimalValue(Text1(12).Text, 2) & "," & DecimalValue(Text1(13).Text, 2) & "," & DecimalValue(Text1(14).Text, 2) & "," & DecimalValue(Text1(15).Text, 2) & "," & DecimalValue(Text1(16).Text, 2) & "," & DecimalValue(Text1(17).Text, 2) & "," & DecimalValue(Text1(18).Text, 2) & "," & DecimalValue(Text1(19).Text, 2) & "," & DecimalValue(Text1(20).Text, 2) & "," & DecimalValue(Text1(21).Text, 2) & "," & DecimalValue(Text1(22).Text, 2) & ")"
    con.Execute sqlstatement
    MsgBox "New Batch ID " & BatchID & " Created!", vbInformation, "DAS Version 4.0"
    ClearForm
End Sub

Private Sub Command2_Click()
    DoEvents
    Dim sqlPrd As String
    Dim BatchID As String
    Dim sqlBrd As String
    Dim appX As New Workbook
    Dim MaxRows As Integer
    
    If (txtTemplate.Text) = "" Then
        MsgBox "New Rate Excel Sheet File Not Selected !", vbExclamation, "DAS Version 4.0"
        Exit Sub
    End If
    If MsgBox("Create New Batch ID ?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbNo Then
        Exit Sub
    End If
    If rec.State = 1 Then rec.Close
    Command2.Enabled = False
    Set appX = GetObject(txtTemplate.Text)
    If UCase(Trim(appX.Sheets(1).Cells(1, 1))) = UCase("Part Number") And UCase(Trim(appX.Sheets(1).Cells(1, 2))) = UCase("PurRate") And UCase(Trim(appX.Sheets(1).Cells(1, 3))) = UCase("SaleRate") And UCase(Trim(appX.Sheets(1).Cells(1, 4))) = UCase("MRPRate") And _
        UCase(Trim(appX.Sheets(1).Cells(1, 5))) = UCase("PurIncDis") And UCase(Trim(appX.Sheets(1).Cells(1, 6))) = UCase("SalIncDis") And UCase(Trim(appX.Sheets(1).Cells(1, 7))) = UCase("PurVAT") And UCase(Trim(appX.Sheets(1).Cells(1, 8))) = UCase("PurCrateRental") And _
        UCase(Trim(appX.Sheets(1).Cells(1, 9))) = UCase("PurCrateVAT") And UCase(Trim(appX.Sheets(1).Cells(1, 10))) = UCase("PurTOTVAT") And UCase(Trim(appX.Sheets(1).Cells(1, 11))) = UCase("PurLPP") And UCase(Trim(appX.Sheets(1).Cells(1, 12))) = UCase("UPurLPP") And _
        UCase(Trim(appX.Sheets(1).Cells(1, 13))) = UCase("SaleVAT") And UCase(Trim(appX.Sheets(1).Cells(1, 14))) = UCase("SaleCrateRental") And UCase(Trim(appX.Sheets(1).Cells(1, 15))) = UCase("SaleCrateVAT") And UCase(Trim(appX.Sheets(1).Cells(1, 16))) = UCase("SaleTOTVAT") And _
        UCase(Trim(appX.Sheets(1).Cells(1, 17))) = UCase("SaleLPP") And UCase(Trim(appX.Sheets(1).Cells(1, 18))) = UCase("USalLPP") And UCase(Trim(appX.Sheets(1).Cells(1, 19))) = UCase("SKU_SBU") And UCase(Trim(appX.Sheets(1).Cells(1, 20))) = UCase("BaseCrateRate") And _
        UCase(Trim(appX.Sheets(1).Cells(1, 21))) = UCase("BaseCrateDis") And UCase(Trim(appX.Sheets(1).Cells(1, 22))) = UCase("BasePurRate") And UCase(Trim(appX.Sheets(1).Cells(1, 23))) = UCase("BaseSaleTax") Then
        Me.MousePointer = vbHourglass
        appX.Sheets(1).Cells(1, 24).Formula = "=counta(a1:a65536)"
        MaxRows = appX.Sheets(1).Cells(1, 24)
        Label7 = MaxRows - 1
        For r = 2 To MaxRows + 1
        If Trim(appX.Sheets(1).Cells(r, 1)) = "" Then
            Exit For
        End If
        sqlPrd = "Select * from prd where "
        If (Option1(0).Value = True) Then
            sqlPrd = sqlPrd & "prd_prn = '" & Trim(appX.Sheets(1).Cells(r, 1)) & "'"
        End If
        If (Option1(1).Value = True) Then
            sqlPrd = sqlPrd & "prd_nme = '" & Trim(appX.Sheets(1).Cells(r, 1)) & "'"
        End If
        If (Option1(2).Value = True) Then
            sqlPrd = sqlPrd & "ALT_IDY = '" & Trim(appX.Sheets(1).Cells(r, 1)) & "'"
        End If
        If rec.State = 1 Then rec.Close
            rec.Open sqlPrd, con, adOpenForwardOnly, adLockReadOnly
            If Not rec.EOF Then
                rec.MoveFirst
                BatchID = "BAT" & T7increment("idy_bat")
                sqlBrd = "insert into brd (prd_idy , bat_idy, bat_nme,  prd_pdr, prd_sdr, prd_mrp, p_inc_dis, s_inc_dis, p_pur_vat, p_crt_rnt, p_crt_vat, p_tot_vat, pur_lpp, u_pur_lpp, s_sal_vat, s_crt_rnt, s_crt_vat, s_tot_vat, sal_lpp, u_sal_lpp, sku_sbu, b_crt_rat, b_crt_dis, b_pur_tax, b_sal_tax) values ("
                sqlBrd = sqlBrd & " '" & rec("prd_idy") & "','" & BatchID & "','" & "MRP : " & DecimalValue(appX.Sheets(1).Cells(r, 4), 2) & "'," & _
                DecimalValue(appX.Sheets(1).Cells(r, 2), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 3), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 4), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 5), 2) & "," & _
                DecimalValue(appX.Sheets(1).Cells(r, 6), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 7), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 8), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 9), 2) & "," & _
                DecimalValue(appX.Sheets(1).Cells(r, 10), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 11), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 12), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 13), 2) & "," & _
                DecimalValue(appX.Sheets(1).Cells(r, 14), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 15), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 16), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 17), 2) & "," & _
                DecimalValue(appX.Sheets(1).Cells(r, 18), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 19), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 20), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 21), 2) & "," & _
                DecimalValue(appX.Sheets(1).Cells(r, 22), 2) & "," & DecimalValue(appX.Sheets(1).Cells(r, 23), 2) & ")"
                con.Execute sqlBrd
            End If
            Label18 = r - 1
            Label18.Refresh
        Next
        Me.MousePointer = vbDefault
        MsgBox "New Batches Created for Product SKU", vbInformation, "DAS Version 4.0"
    Else
        MsgBox "Invalid File Selected for Product,SKU - New Batch Rates Data Import", vbExclamation, "DAS Version 4.0"
    End If
    appX.Close 0
    Set appX = Nothing
    Command2.Enabled = True
    If rec.State = 1 Then rec.Close
End Sub

Private Sub Command3_Click()
    Unload frmNewRateTemplate
End Sub

Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lstprd_nme.SetFocus
    End If
End Sub

Private Sub Command4_Click()
    Dim sqlstatement As String
    If lvBatch.SelectedItem Is Nothing Then
        MsgBox "          Batch  Not Selected            ", vbExclamation, "DAS Version 4.0"
        lstprd_nme.SetFocus
        Exit Sub
    End If
    If MsgBox("Update Batch ID ?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbNo Then
        Exit Sub
    End If
    sqlstatement = "update  brd  set "
    sqlstatement = sqlstatement & " prd_idy = '" & lstprd_idy.Text & "'" & ","
    sqlstatement = sqlstatement & " Bat_nme = " & "'MRP : " & CStr(Text1(3).Text) & "'" & ","
    sqlstatement = sqlstatement & " prd_pdr = " & DecimalValue(Text1(1).Text, 2) & ","
    sqlstatement = sqlstatement & " prd_sdr = " & DecimalValue(Text1(2).Text, 2) & ","
    sqlstatement = sqlstatement & " prd_mrp = " & DecimalValue(Text1(3).Text, 2) & ","
    sqlstatement = sqlstatement & " p_inc_dis = " & DecimalValue(Text1(4).Text, 2) & ","
    sqlstatement = sqlstatement & " s_inc_dis = " & DecimalValue(Text1(5).Text, 2) & ","
    sqlstatement = sqlstatement & " p_pur_vat = " & DecimalValue(Text1(6).Text, 2) & ","
    sqlstatement = sqlstatement & " p_crt_rnt = " & DecimalValue(Text1(7).Text, 2) & ","
    sqlstatement = sqlstatement & " p_crt_vat = " & DecimalValue(Text1(8).Text, 2) & ","
    sqlstatement = sqlstatement & " p_tot_vat = " & DecimalValue(Text1(9).Text, 2) & ","
    sqlstatement = sqlstatement & " pur_lpp = " & DecimalValue(Text1(10).Text, 2) & ","
    sqlstatement = sqlstatement & " u_pur_lpp = " & DecimalValue(Text1(11).Text, 2) & ","
    sqlstatement = sqlstatement & " s_sal_vat = " & DecimalValue(Text1(12).Text, 2) & ","
    sqlstatement = sqlstatement & " s_crt_rnt = " & DecimalValue(Text1(13).Text, 2) & ","
    sqlstatement = sqlstatement & " s_crt_vat = " & DecimalValue(Text1(14).Text, 2) & ","
    sqlstatement = sqlstatement & " s_tot_vat = " & DecimalValue(Text1(15).Text, 2) & ","
    sqlstatement = sqlstatement & " sal_lpp = " & DecimalValue(Text1(16).Text, 2) & ","
    sqlstatement = sqlstatement & " u_sal_lpp = " & DecimalValue(Text1(17).Text, 2) & ","
    sqlstatement = sqlstatement & " sku_sbu = " & DecimalValue(Text1(18).Text, 2) & ","
    sqlstatement = sqlstatement & " b_crt_rat = " & DecimalValue(Text1(19).Text, 2) & ","
    sqlstatement = sqlstatement & " b_crt_dis = " & DecimalValue(Text1(20).Text, 2) & ","
    sqlstatement = sqlstatement & " b_pur_tax = " & DecimalValue(Text1(21).Text, 2) & ","
    sqlstatement = sqlstatement & " b_sal_tax = " & DecimalValue(Text1(22).Text, 2)
    sqlstatement = sqlstatement & " where bat_idy = '" & lvBatch.SelectedItem.SubItems(23) & "'"
    con.Execute sqlstatement
    MsgBox " Batch ID " & lvBatch.SelectedItem.SubItems(23) & " Updated!", vbInformation, "DAS Version 4.0"
    ClearForm
End Sub

Private Sub Command5_Click()
    ClearForm
End Sub


Private Sub ClearForm()
    For i = 0 To Text1.COUNT - 1
        Text1(i).Text = ""
    Next
    lstprd_nme.ListIndex = -1
    lvBatch.ListItems.CLEAR
    txtSearch.Text = ""
    txtTemplate.Text = ""
    lstprd_nme_Click
    txtSearch.SetFocus
'    lstprd_nme.SetFocus
End Sub

Private Function DecimalValue(Value As String, DecimalPlaces As Integer) As String
    If IsNull(Value) Then Value = "0"
    'If Value Is Nothing Then Value = "0"
    If Value = "" Then Value = "0"
    If IsNumeric(Value) = False Then
        Value = "0"
    End If
    Value = CStr(Format(Value, IIf(DecimalPlaces = 3, "##########0.000", "##########0.00")))
    DecimalValue = Value
End Function

Private Function NumberFormat(KeyAscii As Integer, Value As String) As Integer
    If KeyAscii = 46 Then
        If InStr(1, Value, ".") > 0 Then KeyAscii = 0
    Else
        If KeyAscii <> 8 And KeyAscii <> 46 And (KeyAscii < 47 Or KeyAscii > 58) Then KeyAscii = 0
    End If
    NumberFormat = KeyAscii
End Function

Public Function T7increment(FLDVAL As String) As String
'If CHECKTB("idy.dbf") = 1 Then
'    MsgBox UCase(TNAME) & " Missing "
'    Exit Function
'End If

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



