VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RST_VCH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Numbers Edit"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "RST_VCH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5145
      Left            =   0
      TabIndex        =   1
      Top             =   430
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      TabCaption(0)   =   "&File"
      TabPicture(0)   =   "RST_VCH.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label32"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text32"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "&Core"
      TabPicture(1)   =   "RST_VCH.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label28"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label19"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label18"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label17"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label15"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label14"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label13"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label12"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label11"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label10"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label9"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label31"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text20"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Text19"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Text18"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text17"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Text16"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Text15"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Text14"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Text13"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Text12"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Text11"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Text10"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Text9"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Text31"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "&Accounts"
      TabPicture(2)   =   "RST_VCH.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label30"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label29"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label27"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label26"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label25"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label24"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label23"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label22"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label21"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label20"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Text30"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Text29"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Text28"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Text27"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Text26"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Text25"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Text24"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Text23"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Text22"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Text21"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).ControlCount=   20
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2520
         Width           =   825
      End
      Begin VB.TextBox Text31 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68280
         MaxLength       =   5
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   3000
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1080
         Width           =   825
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1560
         Width           =   825
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2040
         Width           =   825
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3000
         Width           =   825
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   3480
         Width           =   825
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3960
         Width           =   825
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4440
         Width           =   825
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1080
         Width           =   825
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1560
         Width           =   825
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2040
         Width           =   825
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2520
         Width           =   825
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3000
         Width           =   825
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -71880
         MaxLength       =   5
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3480
         Width           =   825
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68280
         MaxLength       =   5
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   825
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68280
         MaxLength       =   5
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1080
         Width           =   825
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68280
         MaxLength       =   5
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1560
         Width           =   825
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68280
         MaxLength       =   5
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2040
         Width           =   825
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72120
         MaxLength       =   10
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   840
         Width           =   825
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72120
         MaxLength       =   10
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1320
         Width           =   825
      End
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72120
         MaxLength       =   10
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1800
         Width           =   825
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72120
         MaxLength       =   10
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2280
         Width           =   825
      End
      Begin VB.TextBox Text25 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -72120
         MaxLength       =   10
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2760
         Width           =   825
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68040
         MaxLength       =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   840
         Width           =   825
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68040
         MaxLength       =   10
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1320
         Width           =   825
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68040
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1800
         Width           =   825
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68040
         MaxLength       =   10
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2280
         Width           =   825
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68280
         MaxLength       =   5
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2520
         Width           =   825
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -68040
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2760
         Width           =   825
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Categories"
         Height          =   195
         Left            =   1800
         TabIndex        =   65
         Top             =   2520
         Width           =   750
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Orders"
         Height          =   195
         Left            =   -69030
         TabIndex        =   62
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Group Accounts"
         Height          =   195
         Left            =   1395
         TabIndex        =   60
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Accounts"
         Height          =   195
         Left            =   1875
         TabIndex        =   59
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Locations"
         Height          =   195
         Left            =   1860
         TabIndex        =   58
         Top             =   1560
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Zones"
         Height          =   195
         Left            =   2100
         TabIndex        =   57
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Markets"
         Height          =   195
         Left            =   1980
         TabIndex        =   56
         Top             =   3000
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "SalesMen"
         Height          =   195
         Left            =   1845
         TabIndex        =   55
         Top             =   3480
         Width           =   705
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Product Categories"
         Height          =   195
         Left            =   1200
         TabIndex        =   54
         Top             =   3960
         Width           =   1350
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Products"
         Height          =   195
         Left            =   1920
         TabIndex        =   53
         Top             =   4440
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Purchases Entry"
         Height          =   195
         Left            =   -73350
         TabIndex        =   52
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Opening Stock"
         Height          =   195
         Left            =   -73260
         TabIndex        =   51
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Credit Notes (Stock)"
         Height          =   195
         Left            =   -73620
         TabIndex        =   50
         Top             =   2520
         Width           =   1425
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Debit Notes (Stock)"
         Height          =   195
         Left            =   -73590
         TabIndex        =   49
         Top             =   3000
         Width           =   1395
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Transfer Inward"
         Height          =   195
         Left            =   -73305
         TabIndex        =   48
         Top             =   3480
         Width           =   1110
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Transfer Outward"
         Height          =   195
         Left            =   -69795
         TabIndex        =   47
         Top             =   600
         Width           =   1230
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Free Goods Issued"
         Height          =   195
         Left            =   -69900
         TabIndex        =   46
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Free Goods Received"
         Height          =   195
         Left            =   -70125
         TabIndex        =   45
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Packing Slip"
         Height          =   195
         Left            =   -69450
         TabIndex        =   44
         Top             =   2520
         Width           =   885
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Cash Sales"
         Height          =   195
         Left            =   -72990
         TabIndex        =   43
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Credit Sales"
         Height          =   195
         Left            =   -73035
         TabIndex        =   42
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Payments"
         Height          =   195
         Left            =   -73125
         TabIndex        =   41
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Bank Receipts"
         Height          =   195
         Left            =   -73485
         TabIndex        =   40
         Top             =   2280
         Width           =   1050
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Cash Receipts"
         Height          =   195
         Left            =   -73470
         TabIndex        =   39
         Top             =   2760
         Width           =   1035
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Receipts"
         Height          =   195
         Left            =   -69045
         TabIndex        =   38
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Opening Balance"
         Height          =   195
         Left            =   -69645
         TabIndex        =   37
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Journal Entries"
         Height          =   195
         Left            =   -69450
         TabIndex        =   36
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Credit Notes"
         Height          =   195
         Left            =   -69285
         TabIndex        =   35
         Top             =   2280
         Width           =   870
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Debit Notes"
         Height          =   195
         Left            =   -69255
         TabIndex        =   34
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Bank Payments"
         Height          =   195
         Left            =   -73545
         TabIndex        =   33
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Cash Payments"
         Height          =   195
         Left            =   -73530
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Material Packing"
         Height          =   195
         Left            =   -69750
         TabIndex        =   31
         Top             =   2040
         Width           =   1185
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":0496
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":09DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":0DA2
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":10F6
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":172E
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":1A82
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":1EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":21D6
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":25E2
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":26EE
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":2A42
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RST_VCH.frx":2E56
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "New (Ctrl+N)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s"
            Object.ToolTipText     =   "Save (Ctrl+S)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "c"
            Object.ToolTipText     =   "Cancel (Ctrl+Q)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ed"
            Object.ToolTipText     =   "View (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Category"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "brw"
            Object.ToolTipText     =   "Browse (Ctrl+B)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "d"
            Object.ToolTipText     =   "Delete (Ctrl+D)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "RST_VCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub save()
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If MsgBox("Confirm Values Save ?", vbExclamation + vbYesNo, "Effmcg") = vbYes Then
    rs.Open "select * from idy", con, adOpenKeyset, adLockOptimistic
    rs("idy_grp") = Trim(Val(Text1) & "")
    rs("idy_act") = Trim(Val(Text2) & "")
    rs("idy_LOC") = Trim(Val(Text3) & "")
    rs("idy_zon") = Trim(Val(Text4) & "")
    rs("idy_chn") = Trim(Val(Text32) & "")
    rs("idy_mar") = Trim(Val(Text5) & "")
    rs("idy_smn") = Trim(Val(Text6) & "")
    rs("idy_cat") = Trim(Val(Text7) & "")
    rs("idy_prd") = Trim(Val(Text8) & "")
    rs("idy_cas") = Trim(Val(Text9) & "")
    rs("idy_cri") = Trim(Val(Text10) & "")
    rs("idy_pur") = Trim(Val(Text11) & "")
    rs("idy_ops") = Trim(Val(Text12) & "")
    rs("idy_cns") = Trim(Val(Text13) & "")
    rs("idy_dns") = Trim(Val(Text14) & "")
    rs("idy_tri") = Trim(Val(Text15) & "")
    rs("idy_tro") = Trim(Val(Text16) & "")
    rs("idy_fgi") = Trim(Val(Text17) & "")
    rs("idy_fgr") = Trim(Val(Text18) & "")
    rs("idy_pkg") = Trim(Val(Text19) & "")
    rs("idy_pks") = Trim(Val(Text20) & "")
    rs("idy_bpv") = Trim(Val(Text21) & "")
    rs("idy_cpv") = Trim(Val(Text22) & "")
    rs("idy_pay") = Trim(Val(Text23) & "")
    rs("idy_brv") = Trim(Val(Text24) & "")
    rs("idy_crv") = Trim(Val(Text25) & "")
    rs("idy_rep") = Trim(Val(Text26) & "")
    rs("idy_obn") = Trim(Val(Text27) & "")
    rs("idy_jvn") = Trim(Val(Text28) & "")
    rs("idy_crn") = Trim(Val(Text29) & "")
    rs("idy_drn") = Trim(Val(Text30) & "")
    rs("idy_ord") = Trim(Val(Text31) & "")
    rs.Update
    SSTab1.Tab = 0
    Text1.SetFocus
End If
End Sub

Private Sub cancel()
For A = 1 To RST_VCH.Count - 1
If TypeOf RST_VCH.Controls(A) Is TextBox Then RST_VCH.Controls(A).Text = "0"
Next
SSTab1.Tab = 0
Text1.SetFocus
End Sub
Private Sub Form_Activate()
SSTab1.Tab = 0
Text1.SetFocus
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from idy", con, adOpenKeyset, adLockOptimistic
Text1 = Trim(rs("idy_grp"))
Text2 = Trim(rs("idy_act"))
Text3 = Trim(rs("idy_LOC"))
Text4 = Trim(rs("idy_zon"))
Text32 = Trim(rs("idy_chn"))
Text5 = Trim(rs("idy_mar"))
Text6 = Trim(rs("idy_smn"))
Text7 = Trim(rs("idy_cat"))
Text8 = Trim(rs("idy_prd"))
Text9 = Trim(rs("idy_cas"))
Text10 = Trim(rs("idy_cri"))
Text11 = Trim(rs("idy_pur"))
Text12 = Trim(rs("idy_ops"))
Text13 = Trim(rs("idy_cns"))
Text14 = Trim(rs("idy_dns"))
Text15 = Trim(rs("idy_tri"))
Text16 = Trim(rs("idy_tro"))
Text17 = Trim(rs("idy_fgi"))
Text18 = Trim(rs("idy_fgr"))
Text19 = Trim(rs("idy_pkg"))
Text20 = Trim(rs("idy_pks"))
Text21 = Trim(rs("idy_bpv"))
Text22 = Trim(rs("idy_cpv"))
Text23 = Trim(rs("idy_pay"))
Text24 = Trim(rs("idy_brv"))
Text25 = Trim(rs("idy_crv"))
Text26 = Trim(rs("idy_rep"))
Text27 = Trim(rs("idy_obn"))
Text28 = Trim(rs("idy_jvn"))
Text29 = Trim(rs("idy_crn"))
Text30 = Trim(rs("idy_drn"))
Text31 = Trim(rs("idy_ord"))
End Sub

Private Sub Form_KeyDown(K As Integer, S As Integer)
If K = 120 Then
    Shell "c:\windows\calc.exe"
ElseIf K = 88 And S = 2 Then
    Unload Me
ElseIf K = 83 And S = 2 Then
    save
ElseIf K = 81 And S = 2 Then
    cancel
ElseIf K = 112 Then
    M_HELP
End If
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    If RST_VCH.ActiveControl.name = "Text1" Then
        Unload Me
    Else
        save
    End If
End If
M_MDI.STS_BAR.Panels(1).Text = ""
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    Text1.SetFocus
ElseIf SSTab1.Tab = 1 Then
    Text9.SetFocus
ElseIf SSTab1.Tab = 2 Then
    Text21.SetFocus
End If
End Sub

Private Sub Text1_Change()
If Len(Trim(Text1)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text2_Change()
If Len(Trim(Text2)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text3_Change()
If Len(Trim(Text3)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text4_Change()
If Len(Trim(Text4)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text32_Change()
If Len(Trim(Text32)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text5_Change()
If Len(Trim(Text5)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text6_Change()
If Len(Trim(Text6)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text7_Change()
If Len(Trim(Text7)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text8_Change()
If Len(Trim(Text8)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text9_Change()
If Len(Trim(Text9)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text10_Change()
If Len(Trim(Text10)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text11_Change()
If Len(Trim(Text11)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text12_Change()
If Len(Trim(Text12)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text13_Change()
If Len(Trim(Text13)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text14_Change()
If Len(Trim(Text14)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text15_Change()
If Len(Trim(Text15)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text16_Change()
If Len(Trim(Text16)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text17_Change()
If Len(Trim(Text17)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text18_Change()
If Len(Trim(Text18)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text19_Change()
If Len(Trim(Text19)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text20_Change()
If Len(Trim(Text20)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text21_Change()
If Len(Trim(Text21)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text22_Change()
If Len(Trim(Text22)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text23_Change()
If Len(Trim(Text23)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text24_Change()
If Len(Trim(Text24)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text25_Change()
If Len(Trim(Text25)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text26_Change()
If Len(Trim(Text26)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text27_Change()
If Len(Trim(Text27)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text28_Change()
If Len(Trim(Text28)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text29_Change()
If Len(Trim(Text29)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text30_Change()
If Len(Trim(Text30)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text31_Change()
If Len(Trim(Text31)) = 5 Then SendKeys "{home}+{end}"
End Sub
Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Group Account ID Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Account ID Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Location ID Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text4_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Zone ID Number"
SendKeys "{home}+{end}"
End Sub

Private Sub Text32_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Channel ID Number"
SendKeys "{home}+{end}"
End Sub

Private Sub Text5_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Market ID Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text6_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Salesmen ID Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text7_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Category ID Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text8_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Product ID Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text9_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Cash Sales Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text10_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Credit Sales Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text11_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Purchase Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text12_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Opening Stock Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text13_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Stock Credit Notes Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text14_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Stock Debit Notes Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text15_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Transfer Inwards Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text16_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Transfer Outwards Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text17_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Free Goods Issued Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text18_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Free Goods Received Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text19_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Material Packing Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text20_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Packing Slip Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text21_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Bank Payments Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text22_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Cash Payments Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text23_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Payments Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text24_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Bank Receipts Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text25_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Cash Receipts Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text26_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Receipts Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text27_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Opening Balance Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text28_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Journal Entries Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text29_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Credit Notes Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text30_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Debit Notes Number"
SendKeys "{home}+{end}"
End Sub
Private Sub Text31_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Edit Orders Number"
SendKeys "{home}+{end}"
End Sub

Private Sub Text1_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Text19_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Text2_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text3.SetFocus
ElseIf K = 38 Then
    Text1.SetFocus
End If
End Sub
Private Sub Text3_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text4.SetFocus
ElseIf K = 38 Then
    Text2.SetFocus
End If
End Sub
Private Sub Text4_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text32.SetFocus
ElseIf K = 38 Then
    Text3.SetFocus
End If
End Sub
Private Sub Text32_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text5.SetFocus
ElseIf K = 38 Then
    Text4.SetFocus
End If
End Sub

Private Sub Text5_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text6.SetFocus
ElseIf K = 38 Then
    Text4.SetFocus
End If
End Sub
Private Sub Text6_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text7.SetFocus
ElseIf K = 38 Then
    Text5.SetFocus
End If
End Sub
Private Sub Text7_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text8.SetFocus
ElseIf K = 38 Then
    Text6.SetFocus
End If
End Sub
Private Sub Text8_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    SSTab1.Tab = 1
    Text9.SetFocus
ElseIf K = 38 Then
    Text7.SetFocus
End If
End Sub
Private Sub Text9_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text10.SetFocus
ElseIf K = 38 Then
    SSTab1.Tab = 0
    Text8.SetFocus
End If
End Sub
Private Sub Text10_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text11.SetFocus
ElseIf K = 38 Then
    Text9.SetFocus
End If
End Sub
Private Sub Text11_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text12.SetFocus
ElseIf K = 38 Then
    Text10.SetFocus
End If
End Sub
Private Sub Text12_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text13.SetFocus
ElseIf K = 38 Then
    Text11.SetFocus
End If
End Sub
Private Sub Text13_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text14.SetFocus
ElseIf K = 38 Then
    Text12.SetFocus
End If
End Sub
Private Sub Text14_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text15.SetFocus
ElseIf K = 38 Then
    Text13.SetFocus
End If
End Sub
Private Sub Text15_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text16.SetFocus
ElseIf K = 38 Then
    Text14.SetFocus
End If
End Sub
Private Sub Text16_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text17.SetFocus
ElseIf K = 38 Then
    Text15.SetFocus
End If
End Sub
Private Sub Text17_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text18.SetFocus
ElseIf K = 38 Then
    Text16.SetFocus
End If
End Sub
Private Sub Text18_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text19.SetFocus
ElseIf K = 38 Then
    Text17.SetFocus
End If
End Sub
Private Sub Text19_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text20.SetFocus
ElseIf K = 38 Then
    Text18.SetFocus
End If
End Sub
Private Sub Text20_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text31.SetFocus
ElseIf K = 38 Then
    Text19.SetFocus
End If
End Sub

Private Sub Text31_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    SSTab1.Tab = 2
    Text21.SetFocus
ElseIf K = 38 Then
    Text20.SetFocus
End If
End Sub

Private Sub Text21_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text22.SetFocus
ElseIf K = 38 Then
    SSTab1.Tab = 1
    Text31.SetFocus
End If
End Sub
Private Sub Text22_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text23.SetFocus
ElseIf K = 38 Then
    Text21.SetFocus
End If
End Sub
Private Sub Text23_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text24.SetFocus
ElseIf K = 38 Then
    Text22.SetFocus
End If
End Sub
Private Sub Text24_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text25.SetFocus
ElseIf K = 38 Then
    Text23.SetFocus
End If
End Sub
Private Sub Text25_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text26.SetFocus
ElseIf K = 38 Then
    Text24.SetFocus
End If
End Sub
Private Sub Text26_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text27.SetFocus
ElseIf K = 38 Then
    Text25.SetFocus
End If
End Sub
Private Sub Text27_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text28.SetFocus
ElseIf K = 38 Then
    Text26.SetFocus
End If
End Sub
Private Sub Text28_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text29.SetFocus
ElseIf K = 38 Then
    Text27.SetFocus
End If
End Sub
Private Sub Text29_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    Text30.SetFocus
ElseIf K = 38 Then
    Text28.SetFocus
End If
End Sub
Private Sub Text30_KeyDown(K As Integer, Shift As Integer)
If K = 13 Or K = 40 Then
    save
ElseIf K = 38 Then
    Text29.SetFocus
End If
End Sub
Private Sub Text2_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text3_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text4_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text32_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text5_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text6_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text7_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text8_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text9_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text10_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text11_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text12_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text13_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text14_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text15_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text16_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text17_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text18_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub

Private Sub Text20_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text21_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text31_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text22_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text23_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text24_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text25_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text26_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text27_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text28_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text29_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Text30_KeyPress(K As Integer)
If K > 47 And K < 58 Or K = 8 Then
Else
K = 0
End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "s" Then
    save
ElseIf Button.KEY = "c" Then
    cancel
End If
End Sub


