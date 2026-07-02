VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form M_FRM_NBC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Batch Addition"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "M_FRM_NBC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "sc&Hemes Copy"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1480
      TabIndex        =   51
      Top             =   3165
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Validate Batch"
      Height          =   375
      Left            =   1480
      TabIndex        =   50
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tax Incentive Amount"
      Height          =   1095
      Left            =   2858
      TabIndex        =   45
      Top             =   2400
      Width           =   2295
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   47
         Text            =   "0.00"
         Top             =   580
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   46
         Text            =   "0.00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Sale Inc'tve"
         Height          =   195
         Left            =   45
         TabIndex        =   49
         Top             =   648
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Pur Inc'tve"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   308
         Width           =   765
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Pur Inc'tve Calculation Utility"
      Height          =   2895
      Left            =   323
      TabIndex        =   30
      Top             =   3840
      Width           =   3600
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   36
         Text            =   "0"
         Top             =   580
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   35
         Text            =   "0.00"
         Top             =   930
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   1270
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   1620
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   1965
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   2310
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "SKU Total VAT Amount"
         Height          =   195
         Left            =   375
         TabIndex        =   44
         Top             =   998
         Width           =   1680
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Crate Rental Incl' Tax"
         Height          =   195
         Left            =   525
         TabIndex        =   43
         Top             =   2033
         Width           =   1530
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Basic Rate Per CASE"
         Height          =   195
         Left            =   525
         TabIndex        =   42
         Top             =   308
         Width           =   1530
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Total Quantity"
         Height          =   195
         Left            =   1065
         TabIndex        =   41
         Top             =   648
         Width           =   990
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Tax %"
         Height          =   195
         Left            =   900
         TabIndex        =   40
         Top             =   1338
         Width           =   1155
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Inc'tve Amount Per CASE"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   1688
         Width           =   1815
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Landing Rate Per CASE"
         Height          =   195
         Left            =   345
         TabIndex        =   38
         Top             =   2378
         Width           =   1710
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Sale Basic Rate Calculation Utility"
      Height          =   2895
      Left            =   3953
      TabIndex        =   19
      Top             =   3840
      Width           =   3735
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   1305
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   1650
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   1995
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   2340
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Sale Landing Rate Per CASE"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   308
         Width           =   2070
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Sales Tax%"
         Height          =   195
         Left            =   1365
         TabIndex        =   28
         Top             =   1365
         Width           =   825
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Inc'tve Amount Per CASE"
         Height          =   195
         Left            =   375
         TabIndex        =   27
         Top             =   1725
         Width           =   1815
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Crate Rental Incl' Tax"
         Height          =   195
         Left            =   660
         TabIndex        =   26
         Top             =   2070
         Width           =   1530
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Basic Rate Per CASE"
         Height          =   195
         Left            =   660
         TabIndex        =   25
         Top             =   2415
         Width           =   1530
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Utility"
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   3165
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5287
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5287
      TabIndex        =   16
      Top             =   2775
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rate Per SKU CASE"
      Height          =   1695
      Left            =   1212
      TabIndex        =   9
      Top             =   600
      Width           =   2775
      Begin VB.TextBox PUR_RAT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox MRP_RAT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox SAL_RAT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Rate"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   428
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MRP"
         Height          =   195
         Left            =   885
         TabIndex        =   14
         Top             =   1148
         Width           =   360
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Rate"
         Height          =   195
         Left            =   540
         TabIndex        =   13
         Top             =   788
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SKU Mfg-BBD-Exp"
      Height          =   1695
      Left            =   4024
      TabIndex        =   2
      Top             =   600
      Width           =   2775
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74383363
         CurrentDate     =   37632
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   74383363
         CurrentDate     =   37632
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mfg.Date"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   428
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BBD-Exp.Date"
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   1148
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Shelf Life Days"
         Height          =   195
         Left            =   75
         TabIndex        =   6
         Top             =   788
         Width           =   1065
      End
   End
   Begin VB.TextBox Bat_nme 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3491
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Batch Name"
      Height          =   195
      Left            =   2104
      TabIndex        =   1
      Top             =   195
      Width           =   1260
   End
End
Attribute VB_Name = "M_FRM_NBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public INC_BAT_IDY

Private Sub Bat_nme_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(Bat_nme)) = 0 Then
        MsgBox "Batch Number or ID Required!", vbExclamation, "DAS Version 4.0"
        Bat_nme.SetFocus
    Else
        PUR_RAT.SetFocus
    End If
End If
End Sub

Private Sub Bat_nme_Keypress(K As Integer)
K = CHGCASE(K)
End Sub

Private Sub Command1_Click()
If MsgBox("Confirm Batch Creation and Save?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    Bat_nme.SetFocus
    Exit Sub
Else
    CheckShutDownStatus = GetValue("DIR_MSG ", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'")
    If Trim(CheckShutDownStatus) = "IMPROPER SHUTDOWN" Then
        MsgBox ("Please reset the Improper Shutdown flag in System-Shutdown Status"), vbCritical, "DAS Version 4.0"
        Bat_nme.SetFocus
        Exit Sub
    End If
End If

If Len(Trim(Bat_nme)) = 0 Then
    MsgBox "Batch Number or ID Required!", vbExclamation, "DAS Version 4.0"
    Bat_nme.SetFocus
    Exit Sub
End If

If DateDiff("d", DTPicker1, DTPicker2) < 0 Then
    MsgBox "Invalid Manufacturing and BBD-Expiry Dates", vbExclamation, "DAS Version 4.0"
    DTPicker2.SetFocus
    Exit Sub
End If

If Val(Text1) >= 1500 Then
    MsgBox "Invalid Shelf Life Days for a Product, SKU", vbCritical, "DAS Version 4.0"
    Text1.SetFocus
    Exit Sub
End If

If Val(Text2) < 0 Then
    MsgBox "Invalid Purchase Incentive Amount", vbCritical, "DAS Version 4.0"
    Text2.SetFocus
    Exit Sub
End If

If Val(Text3) < 0 Then
    MsgBox "Invalid Sale Incentive Amount", vbCritical, "DAS Version 4.0"
    Text3.SetFocus
    Exit Sub
End If

If Val(PUR_RAT) < 0 Then
    MsgBox "Invalid Batch Purchase Rate per SKU Case", vbCritical, "DAS Version 4.0"
    PUR_RAT.SetFocus
    Exit Sub
End If

If Val(SAL_RAT) < 0 Then
    MsgBox "Invalid Batch Sale Rate per SKU Case", vbCritical, "DAS Version 4.0"
    SAL_RAT.SetFocus
    Exit Sub
End If

If Val(MRP_RAT) < 0 Then
    MsgBox "Invalid MRP per SKU Case", vbCritical, "DAS Version 4.0"
    MRP_RAT.SetFocus
    Exit Sub
End If

Dim RSBRD As New ADODB.Recordset
Dim SKU_RS As New ADODB.Recordset
If Me.Tag = "PUR" Then
    PRODID = Trim(M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 0))
ElseIf Me.Tag = "OPS" Then
    PRODID = Trim(M_FRM_OPS.MS.TextMatrix(M_FRM_OPS.MS.Row, 0))
End If
If RSBRD.State = 1 Then RSBRD.Close
RSBRD.Open "select * from brd WHERE prd_idy like '" & Trim(PRODID) & "' and bat_nme like '" & Trim(Bat_nme) & "' and prd_pdr =" & Val(PUR_RAT) & " and prd_sdr = " & Val(SAL_RAT) & " and prd_mrp =" & Val(MRP_RAT) & " and brd_dom = ctod('" & DTPicker1 & "') and brd_doe = ctod('" & DTPicker2 & "')", con, adOpenKeyset, adLockPessimistic
If RSBRD.RecordCount = 0 Then
    RSBRD.AddNew
    RSBRD!BAT_IDY = "BAT" & T7increment("IDY_BAT")
    INC_BAT_IDY = RSBRD!BAT_IDY
    RSBRD!Bat_nme = Trim(Bat_nme)
    RSBRD!PRD_PDR = Val(PUR_RAT)
    RSBRD!PRD_MRP = Val(MRP_RAT)
    RSBRD!PRD_SDR = Val(SAL_RAT)
    RSBRD!prd_idy = PRODID
    RSBRD!bRD_dom = DTPicker1
    RSBRD!brd_doe = DTPicker2
    RSBRD!prd_qty = 0#
    RSBRD!P_INC_DIS = 0#
    RSBRD!S_INC_DIS = 0#
    RSBRD!B_PUR_TAX = 0#
    RSBRD!B_SAL_TAX = 0#
    RSBRD!B_CRT_RAT = 0#
    RSBRD!B_CRT_DIS = 0#
    RSBRD!SKU_SBU = 0#
    RSBRD!P_PUR_VAT = 0#
    RSBRD!P_CRT_RNT = 0#
    RSBRD!P_CRT_VAT = 0#
    RSBRD!P_TOT_VAT = 0#
    RSBRD!U_PUR_LPP = 0#
    RSBRD!PUR_LPP = 0#
    RSBRD!S_SAL_VAT = 0#
    RSBRD!S_CRT_RNT = 0#
    RSBRD!S_CRT_VAT = 0#
    RSBRD!S_TOT_VAT = 0#
    RSBRD!U_SAL_LPP = 0#
    RSBRD!SAL_LPP = 0#
    RSBRD.Update
    
    INC_CHK = GetValue("dir_msg ", "dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'INC'")
    If Trim(INC_CHK) = "T" Then
        RSBRD!P_INC_DIS = Val(Text2)
        RSBRD!S_INC_DIS = Val(Text3)
    Else
        MsgBox ("Enable Incentive Location Identification - Incentive Values Not Saved!"), vbCritical, "DAS Version 4.0"
    End If
    If SKU_RS.State = 1 Then SKU_RS.Close
    SKU_RS.Open "SELECT * FROM PRD WHERE PRD_IDY LIKE '" & Trim(PRODID) & "'", con, adOpenKeyset, adLockPessimistic
    If SKU_RS.RecordCount > 0 Then
        M_PUR_LPP = 0
        M_SAL_LPP = 0
        If IsNull(SKU_RS("REO_LEV")) = True Then
            SKU_RS("REO_LEV") = 0
            SKU_RS.Update
        End If
        If IsNull(SKU_RS("AVG_RAT")) = True Then
            SKU_RS("AVG_RAT") = 0
            SKU_RS.Update
        End If
        If IsNull(SKU_RS("BAS_RAT")) = True Then
            SKU_RS("BAS_RAT") = 0
            SKU_RS.Update
        End If
        If IsNull(SKU_RS("LST_RAT")) = True Then
            SKU_RS("LST_RAT") = 0
            SKU_RS.Update
        End If
        If IsNull(SKU_RS("LST_DIS")) = True Then
            SKU_RS("LST_DIS") = 0
            SKU_RS.Update
        End If
        If IsNull(SKU_RS("LST_CST")) = True Then
            SKU_RS("LST_CST") = 0
            SKU_RS.Update
        End If
        If IsNull(SKU_RS("ADD_DIS")) = True Then
            SKU_RS("ADD_DIS") = 0
            SKU_RS.Update
        End If
        If IsNull(SKU_RS("CAS_DIS")) = True Then
            SKU_RS("CAS_DIS") = 0
            SKU_RS.Update
        End If
        If IsNull(SKU_RS("EXC_ADD")) = True Then
            SKU_RS("EXC_ADD") = 0
            SKU_RS.Update
        End If
        If IsNull(SKU_RS("BRW_SEQ")) = True Then
            SKU_RS("BRW_SEQ") = 0
            SKU_RS.Update
        End If
        If IsNull(SKU_RS("TAX_TYP")) = True Then
            SKU_RS("TAX_TYP") = ""
            SKU_RS.Update
        End If
        
        RSBRD!B_PUR_TAX = Val(SKU_RS("PUR_TAX"))
        RSBRD!B_SAL_TAX = Val(SKU_RS("SAL_TAX"))
        RSBRD!B_CRT_RAT = Val(SKU_RS("LST_RAT"))
        RSBRD!B_CRT_DIS = Val(SKU_RS("LST_DIS"))
        RSBRD!SKU_SBU = Val(SKU_RS("SUB_UNT"))
        
        RSBRD!P_PUR_VAT = (((Val(PUR_RAT) - Val(Text2)) * Val(SKU_RS("PUR_TAX"))) / 100)
        RSBRD!P_CRT_RNT = Val(SKU_RS("LST_RAT"))
        RSBRD!P_CRT_VAT = ((Val(SKU_RS("LST_RAT")) * Val(SKU_RS("LST_DIS"))) / 100)
        RSBRD!P_TOT_VAT = (((Val(PUR_RAT) - Val(Text2)) * Val(SKU_RS("PUR_TAX"))) / 100) + ((Val(SKU_RS("LST_RAT")) * Val(SKU_RS("LST_DIS"))) / 100)
        RSBRD!PUR_LPP = Round(Val(PUR_RAT) + (((Val(PUR_RAT) - Val(Text2)) * Val(SKU_RS("PUR_TAX"))) / 100) + Val(SKU_RS("LST_RAT")) + ((Val(SKU_RS("LST_RAT")) * Val(SKU_RS("LST_DIS"))) / 100), 1)
        M_PUR_LPP = Round(Val(PUR_RAT) + (((Val(PUR_RAT) - Val(Text2)) * Val(SKU_RS("PUR_TAX"))) / 100) + Val(SKU_RS("LST_RAT")) + ((Val(SKU_RS("LST_RAT")) * Val(SKU_RS("LST_DIS"))) / 100), 1)
'        RSBRD!U_PUR_LPP = (Val(PUR_RAT) + (((Val(PUR_RAT) - Val(Text2)) * Val(SKU_RS("PUR_TAX"))) / 100) + Val(SKU_RS("LST_RAT")) + ((Val(SKU_RS("LST_RAT")) * Val(SKU_RS("LST_DIS"))) / 100)) / Val(SKU_RS("SUB_UNT"))
        RSBRD!U_PUR_LPP = Round(Val(Val(M_PUR_LPP) / Val(SKU_RS("SUB_UNT"))), 2)
        
        RSBRD!S_SAL_VAT = (((Val(SAL_RAT) - Val(Text3)) * Val(SKU_RS("SAL_TAX"))) / 100)
        RSBRD!S_CRT_RNT = Val(SKU_RS("LST_RAT"))
        RSBRD!S_CRT_VAT = ((Val(SKU_RS("LST_RAT")) * Val(SKU_RS("LST_DIS"))) / 100)
        RSBRD!S_TOT_VAT = (((Val(SAL_RAT) - Val(Text3)) * Val(SKU_RS("SAL_TAX"))) / 100) + ((Val(SKU_RS("LST_RAT")) * Val(SKU_RS("LST_DIS"))) / 100)
        RSBRD!SAL_LPP = Round(Val(SAL_RAT) + (((Val(SAL_RAT) - Val(Text3)) * Val(SKU_RS("SAL_TAX"))) / 100) + Val(SKU_RS("LST_RAT")) + ((Val(SKU_RS("LST_RAT")) * Val(SKU_RS("LST_DIS"))) / 100), 1)
        M_SAL_LPP = Round(Val(SAL_RAT) + (((Val(SAL_RAT) - Val(Text3)) * Val(SKU_RS("SAL_TAX"))) / 100) + Val(SKU_RS("LST_RAT")) + ((Val(SKU_RS("LST_RAT")) * Val(SKU_RS("LST_DIS"))) / 100), 1)
'        RSBRD!U_SAL_LPP = (Val(SAL_RAT) + (((Val(SAL_RAT) - Val(Text3)) * Val(SKU_RS("SAL_TAX"))) / 100) + Val(SKU_RS("LST_RAT")) + ((Val(SKU_RS("LST_RAT")) * Val(SKU_RS("LST_DIS"))) / 100)) / Val(SKU_RS("SUB_UNT"))
        RSBRD!U_SAL_LPP = Round(Val(Val(M_SAL_LPP) / Val(SKU_RS("SUB_UNT"))), 2)
        RSBRD.Update
    Else
    
        RSBRD!B_PUR_TAX = 0#
        RSBRD!B_SAL_TAX = 0#
        RSBRD!B_CRT_RAT = 0#
        RSBRD!B_CRT_DIS = 0#
        RSBRD!SKU_SBU = 0#
        RSBRD!P_PUR_VAT = 0#
        RSBRD!P_CRT_RNT = 0#
        RSBRD!P_CRT_VAT = 0#
        RSBRD!P_TOT_VAT = 0#
        RSBRD!PUR_LPP = 0#
        RSBRD!U_PUR_LPP = 0#
        RSBRD!S_SAL_VAT = 0#
        RSBRD!S_CRT_RNT = 0#
        RSBRD!S_CRT_VAT = 0#
        RSBRD!S_TOT_VAT = 0#
        RSBRD!SAL_LPP = 0#
        RSBRD!U_SAL_LPP = 0#
        RSBRD.Update
    End If
   
    If Me.Tag = "PUR" Then
        M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 2) = convert(Val(PUR_RAT))
        M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 8) = convert(Val(MRP_RAT))
        M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 9) = convert(Val(SAL_RAT))
        M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 13) = Trim(RSBRD!BAT_IDY)
        M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 16) = Format(DTPicker1, "dd/MM/YYYY")
        M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 17) = Format(DTPicker2, "dd/MM/YYYY")
    ElseIf Me.Tag = "OPS" Then
        With M_FRM_OPS
        .MS.TextMatrix(.MS.Row, 2) = convert(Val(PUR_RAT))
        .MS.TextMatrix(.MS.Row, 8) = convert(Val(MRP_RAT))
        .MS.TextMatrix(.MS.Row, 9) = convert(Val(SAL_RAT))
        .MS.TextMatrix(.MS.Row, 13) = Trim(RSBRD!BAT_IDY)
        .MS.TextMatrix(.MS.Row, 16) = Format(DTPicker1, "dd/MM/YYYY")
        .MS.TextMatrix(.MS.Row, 17) = Format(DTPicker2, "dd/MM/YYYY")
        End With
    End If
End If
M_FRM_PUR.M_FLX_AMT.Visible = False
'RSBRD.Close
'SKU_RS.Close
Set RSBRD = Nothing
Set SKU_RS = Nothing
If MsgBox("Include Batch Created in Valid Schemes - BASE or Scheme Definition Part?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
    Scheme_Cascade
End If
If MsgBox("Include Batch Created in Valid Schemes - Free or Scheme Applicability Part?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
    Free_Scheme_Cascade
End If
If MsgBox("Include Batch Created in HVO Schemes - BASE or Scheme Definition Part?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
    HVO_Cascade
End If
'Command5.SetFocus
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Me.Height = 7290 Then
    Me.Height = 4155
Else
    Me.Height = 7290
End If
Text4.SetFocus
End Sub

Private Sub Command4_Click()
Dim BRD_VRF As New ADODB.Recordset
If BRD_VRF.State = 1 Then BRD_VRF.Close
BRD_VRF.Open "SELECT * FROM BRD", con, adOpenKeyset, adLockPessimistic
If BRD_VRF.RecordCount > 0 Then BRD_VRF.MoveFirst
While Not BRD_VRF.EOF
    If IsNull(BRD_VRF!P_INC_DIS) = True Then
        BRD_VRF!P_INC_DIS = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!S_INC_DIS) = True Then
        BRD_VRF!S_INC_DIS = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!B_PUR_TAX) = True Then
        BRD_VRF!B_PUR_TAX = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!B_SAL_TAX) = True Then
        BRD_VRF!B_SAL_TAX = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!B_CRT_RAT) = True Then
        BRD_VRF!B_CRT_RAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!B_CRT_DIS) = True Then
        BRD_VRF!B_CRT_DIS = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!SKU_SBU) = True Then
        BRD_VRF!SKU_SBU = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!P_PUR_VAT) = True Then
        BRD_VRF!P_PUR_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!P_CRT_RNT) = True Then
        BRD_VRF!P_CRT_RNT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!P_CRT_VAT) = True Then
        BRD_VRF!P_CRT_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!P_TOT_VAT) = True Then
        BRD_VRF!P_TOT_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!U_PUR_LPP) = True Then
        BRD_VRF!U_PUR_LPP = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!PUR_LPP) = True Then
        BRD_VRF!PUR_LPP = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!S_SAL_VAT) = True Then
        BRD_VRF!S_SAL_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!S_CRT_RNT) = True Then
        BRD_VRF!S_CRT_RNT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!S_CRT_VAT) = True Then
        BRD_VRF!S_CRT_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!S_TOT_VAT) = True Then
        BRD_VRF!S_TOT_VAT = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!U_SAL_LPP) = True Then
        BRD_VRF!U_SAL_LPP = 0#
        BRD_VRF.Update
    End If
    If IsNull(BRD_VRF!SAL_LPP) = True Then
        BRD_VRF!SAL_LPP = 0#
        BRD_VRF.Update
    End If
    BRD_VRF.MoveNext
Wend
'If BRD_VRF.State = 1 Then BRD_VRF.Close
MsgBox "Batch Validation Completed!", vbInformation, "DAS Version 4.0"
End Sub

Private Sub Command5_Click()
If MsgBox("Cascade Schemes for New Batch Created?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbYes Then
    Scheme_Cascade
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Text1.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If DateDiff("d", DTPicker1, DTPicker2) < 0 Then
    MsgBox "Invalid Manufacturing and BBD-Expiry Dates", vbExclamation, "DAS Version 4.0"
    DTPicker2.SetFocus
Else
Command1.SetFocus
End If
End If
End Sub

Private Sub Form_Activate()
'SETYEAR Me 'Changed by PNR for New Batch Creation Module since Batch can go beyond Financial Year
Dim rsprd As New ADODB.Recordset
Dim BAT_CAS As New ADODB.Recordset

If Me.Tag = "OPS" Then
    PRODID = Trim(M_FRM_OPS.MS.TextMatrix(M_FRM_OPS.MS.Row, 0))
    I_BAT_IDY = Trim(M_FRM_OPS.MS.TextMatrix(M_FRM_OPS.MS.Row, 13))
ElseIf Me.Tag = "PUR" Then
    PRODID = Trim(M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 0))
    I_BAT_IDY = Trim(M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 13))
End If

If rsprd.State = 1 Then rsprd.Close
rsprd.Open "Select * from prd where prd_idy like '" & Trim(PRODID) & "'", con, adOpenKeyset, adLockPessimistic
If rsprd.RecordCount > 0 Then
    B_UNT_RAT = 0
    B_UNT_RAT = Val(rsprd!MRP_RAT) / Val(rsprd!SUB_UNT)
    Bat_nme = "MRP:" & Trim(str(B_UNT_RAT)) & " " & Format(Date, "YY") & Format(Date, "MM") & Format(Date, "dd")
    PUR_RAT = convert(Val(rsprd!PUR_RAT))
    SAL_RAT = convert(Val(rsprd!SAL_RAT))
    MRP_RAT = convert(Val(rsprd!MRP_RAT))
    Text1 = Trim(str(rsprd!TOD_PCG))
    If IsNull(rsprd!PUR_TAX) = False Then
        Text7 = convert(Val(rsprd!PUR_TAX))
    Else
        Text7 = 0#
    End If
    If IsNull(rsprd!SAL_TAX) = False Then
        Text12 = convert(Val(rsprd!SAL_TAX))
    Else
        Text12 = 0#
    End If
    If IsNull(rsprd!LST_RAT) = False Then
        Text9 = convert(Val(rsprd!LST_RAT) + ((Val(rsprd!LST_RAT) * Val(rsprd!LST_DIS)) / 100))
    Else
        Text9 = 0#
    End If
    If IsNull(rsprd!LST_RAT) = False Then
        Text14 = convert(Val(rsprd!LST_RAT) + ((Val(rsprd!LST_RAT) * Val(rsprd!LST_DIS)) / 100))
    Else
        Text14 = 0#
    End If
End If

If BAT_CAS.State = 1 Then BAT_CAS.Close
BAT_CAS.Open "SELECT * FROM BRD WHERE BAT_IDY LIKE '" & Trim(I_BAT_IDY) & "'", con, adOpenKeyset, adLockPessimistic
If BAT_CAS.RecordCount > 0 Then
    PUR_RAT = convert(Val(BAT_CAS!PRD_PDR)) 'Code inserted on Prince feedback to have default rates from batch
    SAL_RAT = convert(Val(BAT_CAS!PRD_SDR))
    MRP_RAT = convert(Val(BAT_CAS!PRD_MRP))
    
    Text2 = convert(Val(BAT_CAS("P_INC_DIS")))
    Text3 = convert(Val(BAT_CAS("S_INC_DIS")))
    If Val(Text7) = 0 Then
        If IsNull(BAT_CAS!B_PUR_TAX) = False Then
            Text7 = convert(Val(BAT_CAS("B_PUR_TAX")))
        Else
            Text7 = 0#
        End If
    End If
    If Val(Text9) = 0 Then
        If IsNull(BAT_CAS!P_CRT_RNT) = False Then
            Text9 = convert(Val(BAT_CAS("P_CRT_RNT")) + Val(BAT_CAS("P_CRT_VAT")))
        Else
            Text9 = 0#
        End If
    End If
    If Val(Text12) = 0 Then
        If IsNull(BAT_CAS!B_SAL_TAX) = False Then
            Text12 = convert(Val(BAT_CAS("B_SAL_TAX")))
        Else
            Text12 = 0#
        End If
    End If
    If IsNull(BAT_CAS!S_INC_DIS) = False Then
        Text13 = convert(Val(BAT_CAS("S_INC_DIS"))) 'Value available only in BRD table
    Else
        Text13 = 0#
    End If
    If Val(Text14) = 0 Then
        If IsNull(BAT_CAS!S_CRT_RNT) = False Then
            Text14 = convert(Val(BAT_CAS("S_CRT_RNT")) + Val(BAT_CAS("S_CRT_VAT")))
        Else
            Text14 = 0#
        End If
    End If
End If

DTPicker1.Value = Date
DTPicker2.Value = Date
If DTPicker2.MinDate < DateAdd("D", Val(Text1), DTPicker1) And DTPicker2.MaxDate > DateAdd("D", Val(Text1), DTPicker1) Then
    DTPicker2 = DateAdd("D", Val(Text1), DTPicker1)
End If
Me.Height = 4155

Bat_nme.SetFocus
End Sub

Private Sub Form_KeyPress(K As Integer)
If K = 27 Then
    Unload Me
End If
End Sub

Private Sub PUR_RAT_GotFocus()
SendKeys "{home}+{end}"
End Sub

Private Sub mrp_RAT_GotFocus()
SendKeys "{home}+{end}"
End Sub
Private Sub sal_RAT_GotFocus()
SendKeys "{home}+{end}"
End Sub

Private Sub PUR_RAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
PUR_RAT = convert(PUR_RAT)
SAL_RAT.SetFocus
End If
End Sub

Private Sub PUR_RAT_keypress(K As Integer)
K = CHKNUM(K)
End Sub
Private Sub MRP_RAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Me.Width = 7635 Then
    DTPicker1.SetFocus
Else
    Command1.SetFocus
End If
MRP_RAT = convert(MRP_RAT)
End If
End Sub

Private Sub MRP_RAT_keypress(K As Integer)
K = CHKNUM(K)
End Sub

Private Sub SAL_RAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SAL_RAT = convert(SAL_RAT)
MRP_RAT.SetFocus
End If
End Sub

Private Sub SAL_RAT_keypress(K As Integer)
K = CHKNUM(K)
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K = 13 Then
DTPicker2.MinDate = "01/01/2000"
DTPicker2.MaxDate = "31/12/2015"
If DTPicker2.MinDate < DateAdd("D", Val(Text1), DTPicker1) And DTPicker2.MaxDate > DateAdd("D", Val(Text1), DTPicker1) Then
    DTPicker2 = DateAdd("D", Val(Text1), DTPicker1)
End If
    Command1.SetFocus
End If
End Sub
Private Sub Text11_GotFocus()
SendKeys "{home}+{end}"
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text11 = convert(Text11)
Text2.SetFocus
End If
End Sub

Private Sub Text11_KeyPress(K As Integer)
K = CHKNUM(K)
Text15 = convert(((Val(Text11) - Val(Text14)) + (((Val(Text13) * Val(Text12)) / 100))) * (100 / (100 + Val(Text12))))
End Sub

Private Sub TEXT2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text2 = convert(Text2)
Text3.SetFocus
End If
End Sub
Private Sub Text2_KeyPress(K As Integer)
K = CHKNUM(K)
End Sub
Private Sub Text3_GotFocus()
SendKeys "{home}+{end}"
End Sub
Private Sub Text2_GotFocus()
SendKeys "{home}+{end}"
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text3 = convert(Text3)
Text13 = convert(Text3)
Command1.SetFocus
End If
End Sub
Private Sub text3_KeyPress(K As Integer)
K = CHKNUM(K)
Text13 = convert(Val(Text3))
End Sub
Private Sub Text4_GotFocus()
SendKeys "{home}+{end}"
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text4 = convert(Text4)
Text5.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(K As Integer)
K = CHKNUM(K)
End Sub

Private Sub Text5_GotFocus()
SendKeys "{home}+{end}"
End Sub

Private Sub TEXT5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text5 = Val(Text5)
Text6.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(K As Integer)
K = CHKNUM(K)
End Sub

Private Sub Text6_GotFocus()
SendKeys "{home}+{end}"
End Sub

Private Sub TEXT6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text6 = convert(Text6)
' Calculation Cycle
    If Val(Text5) > 0 And Val(Text7) > 0 Then
        Text8 = convert(Round(Val(Text4) - ((Text6 / Text5) * 100 / Val(Text7)), 2))
    End If
    Text10 = convert(Round(Val(Text4) + ((Val(Text4) - Val(Text8)) * Val(Text7) / 100) + Val(Text9), 2))
Text11.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(K As Integer)
K = CHKNUM(K)
End Sub
Private Sub Scheme_Cascade()
Dim SCH_RS As New ADODB.Recordset
Dim Chk_RS As New ADODB.Recordset
Dim New_RS As New ADODB.Recordset
If Me.Tag = "PUR" Then
    PRODID = Trim(M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 0))
ElseIf Me.Tag = "OPS" Then
    PRODID = Trim(M_FRM_OPS.MS.TextMatrix(M_FRM_OPS.MS.Row, 0))
End If
SCH_RS.Open "SELECT * FROM SCH WHERE B_PRD_IDY LIKE '" & Trim(PRODID) & "' AND SCH_STS = 'Y' AND SCH_BDT <= ctod('" & Date & "') AND SCH_EDT >= ctod('" & Date & "') AND TXN_TYP = 'BAS'", con, adOpenKeyset, adLockPessimistic
If SCH_RS.RecordCount > 0 Then
    While Not SCH_RS.EOF
        If Chk_RS.State = 1 Then Chk_RS.Close
        Chk_RS.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCH_RS("SCH_IDY")) & "' AND BAT_IDY LIKE '" & Trim(INC_BAT_IDY) & "' AND TXN_TYP = 'BAS'", con, adOpenKeyset, adLockPessimistic
        If Chk_RS.EOF Then
            If New_RS.State = 1 Then New_RS.Close
            New_RS.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCH_RS("SCH_IDY")) & "' AND TXN_TYP = 'BAS' GROUP BY B_PRD_IDY", con, adOpenKeyset, adLockPessimistic
            If New_RS.RecordCount > 0 Then
'                NEW_SCH_IDY = "SCH" & T7increment("IDY_SCH")
                While Not New_RS.EOF
                If Trim(New_RS("B_PRD_IDY")) = Trim(PRODID) Then
                    Chk_RS.AddNew
'                    Chk_RS("SCH_IDY") = Trim(NEW_SCH_IDY)
                    Chk_RS("SCH_IDY") = Trim(New_RS("SCH_IDY"))
                    Chk_RS("SCH_NME") = Trim(New_RS("SCH_NME"))
                    Chk_RS("SCH_TYP") = Trim(New_RS("SCH_TYP"))
                    Chk_RS("F_SCH_TYP") = Trim(New_RS("F_SCH_TYP"))
                    Chk_RS("BAT_IDY") = Trim(INC_BAT_IDY)           'adding only new batch record with the same scheme details as earlier
                    Chk_RS("B_PRD_IDY") = Trim(New_RS("B_PRD_IDY"))
                    Chk_RS("B_PRD_QTY") = (New_RS("B_PRD_QTY"))
                    Chk_RS("E_PRD_QTY") = (New_RS("E_PRD_QTY"))
                    Chk_RS("F_PRD_IDY") = Trim(New_RS("F_PRD_IDY"))
                    Chk_RS("F_PRD_QTY") = Trim(New_RS("F_PRD_QTY"))
                    Chk_RS("B_PRD_AMT") = (New_RS("B_PRD_AMT"))
                    Chk_RS("B_DIS_PCG") = (New_RS("B_DIS_PCG"))
                    Chk_RS("B_DIS_AMT") = (New_RS("B_DIS_AMT"))
                    Chk_RS("B_BIL_AMT") = (New_RS("B_BIL_AMT"))
                    Chk_RS("E_BIL_AMT") = (New_RS("E_BIL_AMT"))
                    Chk_RS("SCH_BDT") = New_RS("SCH_BDT")
                    Chk_RS("SCH_EDT") = New_RS("SCH_EDT")
                    Chk_RS("TXN_TYP") = New_RS("TXN_TYP")
                    Chk_RS("SCH_STS") = New_RS("SCH_sTS")
                    Chk_RS("USR_IDY") = New_RS("USR_IDY")
                    Chk_RS("SYS_DAT") = New_RS("SYS_DAT")
                    Chk_RS("SYS_TIM") = New_RS("SYS_TIM")
                    Chk_RS("PRO_RAT") = New_RS("PRO_RAT")
                    Chk_RS("B_CON_APP") = New_RS("B_CON_APP")
                    Chk_RS("F_CON_APP") = New_RS("F_CON_APP")
                    Chk_RS("GRP_TYP") = New_RS("GRP_TYP")
                    Chk_RS("MER_NUM") = New_RS("MER_NUM")
                    Chk_RS("MER_QTY") = (New_RS("MER_QTY"))
                    Chk_RS("MER_AMT") = (New_RS("MER_AMT"))
                    Chk_RS("CMP_DST") = Trim(New_RS("CMP_DST"))
                    Chk_RS.Update
                End If
                New_RS.MoveNext
                Wend
            End If
        End If
        SCH_RS.MoveNext
    Wend
End If
Set New_RS = Nothing
Set Chk_RS = Nothing
Set SCH_RS = Nothing
MsgBox "Corresponding and Valid Base Schemes Updated by DAS for New Batch Created", vbInformation, "DAS Version 4.0"
End Sub
Private Sub Free_Scheme_Cascade()
Dim SCH_RS As New ADODB.Recordset
Dim Chk_RS As New ADODB.Recordset
Dim New_RS As New ADODB.Recordset
If Me.Tag = "PUR" Then
    PRODID = Trim(M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 0))
ElseIf Me.Tag = "OPS" Then
    PRODID = Trim(M_FRM_OPS.MS.TextMatrix(M_FRM_OPS.MS.Row, 0))
End If
SCH_RS.Open "SELECT * FROM SCH WHERE B_PRD_IDY LIKE '" & Trim(PRODID) & "' AND SCH_STS = 'Y' AND SCH_BDT <= ctod('" & Date & "') AND SCH_EDT >= ctod('" & Date & "') AND TXN_TYP = 'FRE'", con, adOpenKeyset, adLockPessimistic
If SCH_RS.RecordCount > 0 Then
    While Not SCH_RS.EOF
        If Chk_RS.State = 1 Then Chk_RS.Close
        Chk_RS.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCH_RS("SCH_IDY")) & "' AND BAT_IDY LIKE '" & Trim(INC_BAT_IDY) & "' AND TXN_TYP = 'FRE'", con, adOpenKeyset, adLockPessimistic
        If Chk_RS.EOF Then
            If New_RS.State = 1 Then New_RS.Close
            New_RS.Open "SELECT * FROM SCH WHERE SCH_IDY LIKE '" & Trim(SCH_RS("SCH_IDY")) & "' AND TXN_TYP = 'FRE' GROUP BY B_PRD_IDY", con, adOpenKeyset, adLockPessimistic
            If New_RS.RecordCount > 0 Then
'                NEW_SCH_IDY = "SCH" & T7increment("IDY_SCH")
                While Not New_RS.EOF
                If Trim(New_RS("B_PRD_IDY")) = Trim(PRODID) Then
                    Chk_RS.AddNew
'                    Chk_RS("SCH_IDY") = Trim(NEW_SCH_IDY)
                    Chk_RS("SCH_IDY") = Trim(New_RS("SCH_IDY"))
                    Chk_RS("SCH_NME") = Trim(New_RS("SCH_NME"))
                    Chk_RS("SCH_TYP") = Trim(New_RS("SCH_TYP"))
                    Chk_RS("F_SCH_TYP") = Trim(New_RS("F_SCH_TYP"))
                    Chk_RS("BAT_IDY") = Trim(INC_BAT_IDY)           'adding only new batch record with the same scheme details as earlier
                    Chk_RS("B_PRD_IDY") = Trim(New_RS("B_PRD_IDY"))
                    Chk_RS("B_PRD_QTY") = (New_RS("B_PRD_QTY"))
                    Chk_RS("E_PRD_QTY") = (New_RS("E_PRD_QTY"))
                    Chk_RS("F_PRD_IDY") = Trim(New_RS("F_PRD_IDY"))
                    Chk_RS("F_PRD_QTY") = Trim(New_RS("F_PRD_QTY"))
                    Chk_RS("B_PRD_AMT") = (New_RS("B_PRD_AMT"))
                    Chk_RS("B_DIS_PCG") = (New_RS("B_DIS_PCG"))
                    Chk_RS("B_DIS_AMT") = (New_RS("B_DIS_AMT"))
                    Chk_RS("B_BIL_AMT") = (New_RS("B_BIL_AMT"))
                    Chk_RS("E_BIL_AMT") = (New_RS("E_BIL_AMT"))
                    Chk_RS("SCH_BDT") = New_RS("SCH_BDT")
                    Chk_RS("SCH_EDT") = New_RS("SCH_EDT")
                    Chk_RS("TXN_TYP") = New_RS("TXN_TYP")
                    Chk_RS("SCH_STS") = New_RS("SCH_sTS")
                    Chk_RS("USR_IDY") = New_RS("USR_IDY")
                    Chk_RS("SYS_DAT") = New_RS("SYS_DAT")
                    Chk_RS("SYS_TIM") = New_RS("SYS_TIM")
                    Chk_RS("PRO_RAT") = New_RS("PRO_RAT")
                    Chk_RS("B_CON_APP") = New_RS("B_CON_APP")
                    Chk_RS("F_CON_APP") = New_RS("F_CON_APP")
                    Chk_RS("GRP_TYP") = New_RS("GRP_TYP")
                    Chk_RS("MER_NUM") = New_RS("MER_NUM")
                    Chk_RS("MER_QTY") = (New_RS("MER_QTY"))
                    Chk_RS("MER_AMT") = (New_RS("MER_AMT"))
                    Chk_RS("CMP_DST") = Trim(New_RS("CMP_DST"))
                    Chk_RS.Update
                End If
                New_RS.MoveNext
                Wend
            End If
        End If
        SCH_RS.MoveNext
    Wend
End If
Set New_RS = Nothing
Set Chk_RS = Nothing
Set SCH_RS = Nothing
MsgBox "Corresponding and Valid Free Schemes Updated by DAS for New Batch Created", vbInformation, "DAS Version 4.0"
End Sub
Private Sub HVO_Cascade()
Dim HVO_RS As New ADODB.Recordset
Dim Chk_RS As New ADODB.Recordset
Dim New_RS As New ADODB.Recordset
If Me.Tag = "PUR" Then
    PRODID = Trim(M_FRM_PUR.MS.TextMatrix(M_FRM_PUR.MS.Row, 0))
ElseIf Me.Tag = "OPS" Then
    PRODID = Trim(M_FRM_OPS.MS.TextMatrix(M_FRM_OPS.MS.Row, 0))
End If
HVO_RS.Open "SELECT * FROM CSR WHERE B_PRD_IDY LIKE '" & Trim(PRODID) & "' AND CSR_STS = 'Y' ", con, adOpenKeyset, adLockPessimistic
 'AND BEG_DAT <= ctod('" & Date & "') AND END_DAT >= ctod('" & Date & "')"
If HVO_RS.RecordCount > 0 Then
    While Not HVO_RS.EOF
        If Chk_RS.State = 1 Then Chk_RS.Close
        Chk_RS.Open "SELECT * FROM CSR WHERE CSR_IDY LIKE '" & Trim(HVO_RS("CSR_IDY")) & "' AND B_BAT_IDY LIKE '" & Trim(INC_BAT_IDY) & "'", con, adOpenKeyset, adLockPessimistic
        If Chk_RS.EOF Then
            If New_RS.State = 1 Then New_RS.Close
            New_RS.Open "SELECT * FROM CSR WHERE CSR_IDY LIKE '" & Trim(HVO_RS("CSR_IDY")) & "' GROUP BY B_PRD_IDY", con, adOpenKeyset, adLockPessimistic
            If New_RS.RecordCount > 0 Then
'                NEW_SCH_IDY = "SCH" & T7increment("IDY_SCH")
                While Not New_RS.EOF
                If Trim(New_RS("B_PRD_IDY")) = Trim(PRODID) Then
                    Chk_RS.AddNew
                    Chk_RS("CSR_IDY") = Trim(New_RS("CSR_IDY"))
                    Chk_RS("XCSR_IDY") = Trim(New_RS("XCSR_IDY"))
                    Chk_RS("CSR_DAT") = New_RS("CSR_DAT")
                    Chk_RS("SCH_IDY") = New_RS("SCH_IDY")
                    Chk_RS("MER_IDY") = New_RS("MER_IDY")
                    Chk_RS("MER_QTY") = New_RS("MER_QTY")
                    Chk_RS("MER_AMT") = New_RS("MER_AMT")
                    Chk_RS("USG_QTY") = New_RS("USG_QTY")
                    Chk_RS("USG_AMT") = New_RS("USG_AMT")
                    Chk_RS("BAL_QTY") = New_RS("BAL_QTY")
                    Chk_RS("BAL_AMT") = New_RS("BAL_AMT")
                    Chk_RS("CLM_QTY") = New_RS("CLM_QTY")
                    Chk_RS("CLM_AMT") = New_RS("CLM_AMT")
                    Chk_RS("CLM_APP") = New_RS("CLM_APP")
                    Chk_RS("CLM_RCD") = New_RS("CLM_RCD")
                    Chk_RS("CLM_PNG") = New_RS("CLM_PNG")
                    Chk_RS("CLM_TYP") = New_RS("CLM_TYP")
                    Chk_RS("CSR_TYP") = New_RS("CSR_TYP")
                    Chk_RS("ACT_IDY") = New_RS("ACT_IDY")
                    Chk_RS("PGR_IDY") = New_RS("PGR_IDY")
                    Chk_RS("BRD_NME") = New_RS("BRD_NME")
                    Chk_RS("PCK_SIZ") = New_RS("PCK_SIZ")
                    Chk_RS("PCK_TYP") = New_RS("PCK_TYP")
                    Chk_RS("PRD_UOM") = New_RS("PRD_UOM")
                    Chk_RS("B_PRD_IDY") = New_RS("B_PRD_IDY")
                    Chk_RS("B_BAT_IDY") = Trim(INC_BAT_IDY)     'adding only new batch record with the same scheme details as earlier
                    Chk_RS("B_PRD_RAT") = New_RS("B_PRD_RAT")
                    Chk_RS("B_PRD_TAX") = New_RS("B_PRD_TAX")
                    Chk_RS("B_PRD_INC") = New_RS("B_PRD_INC")
                    Chk_RS("B_PRD_LST") = New_RS("B_PRD_LST")
                    Chk_RS("B_PRD_LSR") = New_RS("B_PRD_LSR")
                    Chk_RS("B_PRD_SUT") = New_RS("B_PRD_SUT")
                    Chk_RS("B_PRD_UOM") = New_RS("B_PRD_UOM")
                    Chk_RS("B_PRD_QTY") = New_RS("B_PRD_QTY")
                    Chk_RS("B_PRD_PCS") = New_RS("B_PRD_PCS")
                    Chk_RS("B_PRD_CAS") = New_RS("B_PRD_CAS")
                    Chk_RS("B_BAS_QTY") = New_RS("B_BAS_QTY")
                    Chk_RS("B_BAS_AMT") = New_RS("B_BAS_AMT")
                    Chk_RS("B_INT_TYP") = New_RS("B_INT_TYP")
                    Chk_RS("F_INT_TYP") = New_RS("F_INT_TYP")
                    Chk_RS("F_PRD_IDY") = New_RS("F_PRD_IDY")
                    Chk_RS("F_BAT_IDY") = New_RS("F_BAT_IDY")
                    Chk_RS("F_PRD_RAT") = New_RS("F_PRD_RAT")
                    Chk_RS("F_PRD_TAX") = New_RS("F_PRD_TAX")
                    Chk_RS("F_PRD_INC") = New_RS("F_PRD_INC")
                    Chk_RS("F_PRD_LST") = New_RS("F_PRD_LST")
                    Chk_RS("F_PRD_LSR") = New_RS("F_PRD_LSR")
                    Chk_RS("F_PRD_SUT") = New_RS("F_PRD_SUT")
                    Chk_RS("F_PRD_UOM") = New_RS("F_PRD_UOM")
                    Chk_RS("F_PRD_QTY") = New_RS("F_PRD_QTY")
                    Chk_RS("F_PRD_PCS") = New_RS("F_PRD_PCS")
                    Chk_RS("F_PRD_CAS") = New_RS("F_PRD_CAS")
                    Chk_RS("F_FRE_QTY") = New_RS("F_FRE_QTY")
                    Chk_RS("F_FRE_AMT") = New_RS("F_FRE_AMT")
                    Chk_RS("F_CPN_QTY") = New_RS("F_CPN_QTY")
                    Chk_RS("F_CPN_AMT") = New_RS("F_CPN_AMT")
                    Chk_RS("DIS_TYP") = New_RS("DIS_TYP")
                    Chk_RS("DIS_PCG") = New_RS("DIS_PCG")
                    Chk_RS("DIS_AMT") = New_RS("DIS_AMT")
                    Chk_RS("EXP_AMT") = New_RS("EXP_AMT")
                    Chk_RS("BEG_DAT") = New_RS("BEG_DAT")
                    Chk_RS("END_DAT") = New_RS("END_DAT")
                    Chk_RS("CLS_IDY") = New_RS("CLS_IDY")
                    Chk_RS("LST_IDY") = New_RS("LST_IDY")
                    Chk_RS("PRC_CLS") = New_RS("PRC_CLS")
                    Chk_RS("PRC_LST") = New_RS("PRC_LST")
                    Chk_RS("R_BEG_SLB") = New_RS("R_BEG_SLB")
                    Chk_RS("R_END_SLB") = New_RS("R_END_SLB")
                    Chk_RS("Q_BEG_SLB") = New_RS("Q_BEG_SLB")
                    Chk_RS("Q_END_SLB") = New_RS("Q_END_SLB")
                    Chk_RS("S_BEG_SLB") = New_RS("S_BEG_SLB")
                    Chk_RS("S_END_SLB") = New_RS("S_END_SLB")
                    Chk_RS("L_BEG_SLB") = New_RS("L_BEG_SLB")
                    Chk_RS("L_END_SLB") = New_RS("L_END_SLB")
                    Chk_RS("PRO_RAT") = New_RS("PRO_RAT")
                    Chk_RS("CSR_STS") = New_RS("CSR_STS")
                    Chk_RS("SYS_DAT") = New_RS("SYS_DAT")
                    Chk_RS("SYS_TIM") = New_RS("SYS_TIM")
                    Chk_RS("USR_IDY") = New_RS("USR_IDY")
                    Chk_RS("USR_NME") = New_RS("USR_NME")
                    Chk_RS("CMP_DST") = New_RS("CMP_DST")
                    Chk_RS("CSR_NME") = New_RS("CSR_NME")
                    Chk_RS("CLM_TPS") = New_RS("CLM_TPS")
                    Chk_RS("MAR_IDY") = New_RS("MAR_IDY")
                    Chk_RS("SMN_IDY") = New_RS("SMN_IDY")
                    Chk_RS("ROU_IDY") = New_RS("ROU_IDY")
                    Chk_RS("LOC_IDY") = New_RS("LOC_IDY")
                    Chk_RS("CAT_IDY") = New_RS("CAT_IDY")
                    Chk_RS("OUT_TYP") = New_RS("OUT_TYP")
                    Chk_RS("OUT_CLS") = New_RS("OUT_CLS")
                    Chk_RS("OUT_CAP") = New_RS("OUT_CAP")
                    Chk_RS("SGA_CAP") = New_RS("SGA_CAP")
                    Chk_RS.Update
                End If
                New_RS.MoveNext
                Wend
            End If
        End If
        HVO_RS.MoveNext
    Wend
End If
Set New_RS = Nothing
Set Chk_RS = Nothing
Set HVO_RS = Nothing
MsgBox "Corresponding and Valid HVO Schemes Updated by DAS for New Batch Created", vbInformation, "DAS Version 4.0"
End Sub

