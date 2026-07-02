VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form M_REP_TGT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Report Defaults"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   Icon            =   "M_REP_RADHIKA..frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "M_REP_RADHIKA..frx":0442
      Left            =   840
      List            =   "M_REP_RADHIKA..frx":044C
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox ROU_LST 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "M_REP_RADHIKA..frx":0462
      Left            =   1440
      List            =   "M_REP_RADHIKA..frx":04D2
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate Report"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Target Option"
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
      Begin VB.OptionButton Option4 
         Caption         =   "S&KU"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "&Brand"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Segment"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Categotory"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin MSComCtl2.DTPicker DATE1 
      Height          =   330
      Left            =   840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   24444931
      CurrentDate     =   37099
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Route"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   540
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   195
      Width           =   345
   End
End
Attribute VB_Name = "M_REP_TGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim ary()
Dim aryday()
Dim arywek()
Dim arymth()
If ROU_LST = "All" Then
    If Me.Tag = "PNT" Then
        ROUIDY = "%"
    Else
        ROUIDY = "%"
    End If
Else
    ROUIDY = ROU_LST
End If
CON.Execute "update dir set dir_msg= '" & Trim(Format(DATE1, "dd/mmm/yyyy")) & "' where dir_idy like 'HDR' AND DIR_IDX LIKE 'DAT'"
CON.Execute "update dir set dir_als= '" & Trim(ROU_LST) & "' where dir_idy like 'ROU' AND DIR_IDX LIKE 'IDY'"
Dim TRT As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim WEK As New ADODB.Recordset
CON.Execute "DELETE FROM TRT"
WEK.Open "SELECT * FROM WEK WHERE FRM_DAT <=CTOD('" & DATE1 & "') AND TO_DAT >=CTOD('" & DATE1 & "')", CON, adOpenKeyset, adLockOptimistic
If WEK.RecordCount > 0 Then
    FRMDAT = WEK!frm_dat
    TODAT = WEK!to_dat
    DAYS = DateDiff("D", FRMDAT, TODAT) + 1
    wekidy = Val(WEK!WEK_IDY)
    MTHIDY = Val(WEK!MTH_NOS)
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT MIN(FRM_DAT) FROM WEK WHERE MTH_NOS = " & MTHIDY, CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then MINDAT = rs(0)
End If

If rs.State = 1 Then rs.Close       'DAY VALUE UPDATAION
If list1 = "Volume" Then
    rs.Open "SELECT PRD.PRD_IDY,SUM(inv.prd_uom),grp_cat,GRP_CLS,BRD_NME FROM INV,prd,pgr WHERE inv.prd_idy like prd.prd_idy and prd.grp_idy like pgr.grp_idy and INV_IDY IN(SELECT INV_IDY FROM INL WHERE INV_DAT = CTOD('" & DATE1 & "') AND TXN_TYP LIKE 'SAL' AND ROU_IDY LIKE '" & Trim(ROUIDY) & "')AND INV.STK_TYP IN('Salable','Free') GROUP BY INV.PRD_IDY", CON, adOpenKeyset, adLockOptimistic
Else
    rs.Open "SELECT PRD.PRD_IDY,SUM(tot_pcs),grp_cat,GRP_CLS,BRD_NME,inv.sub_unt FROM INV,prd,pgr WHERE inv.prd_idy like prd.prd_idy and prd.grp_idy like pgr.grp_idy and INV_IDY IN(SELECT INV_IDY FROM INL WHERE INV_DAT = CTOD('" & DATE1 & "') AND TXN_TYP LIKE 'SAL' AND ROU_IDY LIKE '" & Trim(ROUIDY) & "')AND INV.STK_TYP IN('Salable','Free') GROUP BY INV.PRD_IDY", CON, adOpenKeyset, adLockOptimistic
End If
If rs.RecordCount > 0 Then
    If TRT.State = 1 Then TRT.Close
    TRT.Open "SELECT * FROM TRT", CON, adOpenKeyset, adLockOptimistic
    For A = 1 To rs.RecordCount
        TRT.AddNew
        TRT!prd_idy = Trim(rs(0))
        TRT!prd_cat = Trim(rs(2))
        TRT!prd_cls = Trim(rs(3))
        TRT!BRD_NME = Trim(rs(4))
        If Trim(list1) = "Volume" Then
            TRT!a_DAY_stk = Val(rs(1) & "")
        Else
            TRT!a_DAY_stk = STM(Val(rs(1) & ""), Val(rs(5)))
        End If
        TRT!a_wek_stk = 0
        TRT!a_mth_stk = 0
        TRT!T_DAY_STK = 0
        TRT!T_WEK_STK = 0
        TRT!T_MTH_STK = 0
        TRT!MTH_DYS = DateDiff("D", MINDAT, TODAT) + 1
        TRT!WEK_DYS = Val(DAYS)
        TRT!CMP_WEK = DateDiff("D", FRMDAT, DATE1) + 1
        TRT!CMP_MTH = DateDiff("D", MINDAT, DATE1) + 1
        TRT.Update
        rs.MoveNext
    Next
End If

If rs.State = 1 Then rs.Close 'WEEK VALUE UPDATAION
If Trim(list1) = "Volume" Then
    rs.Open "SELECT PRD.PRD_IDY,SUM(INV.PRD_UOM),grp_cat,GRP_CLS,BRD_NME FROM INV,prd,pgr WHERE inv.prd_idy like prd.prd_idy and prd.grp_idy like pgr.grp_idy and INV_IDY IN(SELECT INV_IDY FROM INL WHERE INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & FRMDAT & "') AND TXN_TYP LIKE 'SAL' AND ROU_IDY LIKE '" & Trim(ROUIDY) & "')AND INV.STK_TYP IN('Salable','Free') GROUP BY INV.PRD_IDY", CON, adOpenKeyset, adLockOptimistic
Else
    rs.Open "SELECT PRD.PRD_IDY,SUM(tot_pcs),grp_cat,GRP_CLS,BRD_NME,inv.sub_unt FROM INV,prd,pgr WHERE inv.prd_idy like prd.prd_idy and prd.grp_idy like pgr.grp_idy and INV_IDY IN(SELECT INV_IDY FROM INL WHERE INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & FRMDAT & "') AND TXN_TYP LIKE 'SAL' AND ROU_IDY LIKE '" & Trim(ROUIDY) & "')AND INV.STK_TYP IN('Salable','Free') GROUP BY INV.PRD_IDY", CON, adOpenKeyset, adLockOptimistic
End If
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
    If Trim(list1) = "Volume" Then
        CON.Execute "UPDATE trt SET a_wek_stk = " & Val(rs(1) & "") & " WHERE PRD_IDY LIKE '" & Trim(rs(0)) & "'", d
    Else
        CON.Execute "UPDATE trt SET a_wek_stk = " & STM(Val(rs(1) & ""), Val(rs(5) & "")) & " WHERE PRD_IDY LIKE '" & Trim(rs(0)) & "'", d
    End If
    If d = 0 Then
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT", CON, adOpenKeyset, adLockOptimistic
        TRT.AddNew
        TRT!prd_idy = Trim(rs(0))
        TRT!prd_cat = Trim(rs(2))
        TRT!prd_cls = Trim(rs(3))
        TRT!BRD_NME = Trim(rs(4))
        If Trim(list1) = "Volume" Then
            TRT!a_wek_stk = Val(rs(1) & "")
        Else
            TRT!a_wek_stk = STM(Val(rs(1) & ""), Val(rs(5) & ""))
        End If
        TRT!a_DAY_stk = 0
        TRT!a_mth_stk = 0
        TRT!T_DAY_STK = 0
        TRT!T_WEK_STK = 0
        TRT!T_MTH_STK = 0
        TRT!MTH_DYS = DateDiff("D", MINDAT, TODAT) + 1
        TRT!WEK_DYS = Val(DAYS)
        TRT!CMP_WEK = DateDiff("D", FRMDAT, DATE1) + 1
        TRT!CMP_MTH = DateDiff("D", MINDAT, DATE1) + 1
        TRT.Update
    End If
    rs.MoveNext
    Next
End If

If rs.State = 1 Then rs.Close 'MTH VALUE UPDATAION
If list1 = "Volume" Then
    rs.Open "SELECT PRD.PRD_IDY,SUM(INV.PRD_UOM),grp_cat,GRP_CLS,BRD_NME FROM INV,prd,pgr WHERE inv.prd_idy like prd.prd_idy and prd.grp_idy like pgr.grp_idy and INV_IDY IN(SELECT INV_IDY FROM INL WHERE INV_DAT <= CTOD('" & DATE1 & "')  AND INV_DAT >=CTOD('" & MINDAT & "') AND TXN_TYP LIKE 'SAL' AND ROU_IDY LIKE '" & Trim(ROUIDY) & "')AND INV.STK_TYP IN('Salable','Free') GROUP BY INV.PRD_IDY", CON, adOpenKeyset, adLockOptimistic
Else
    rs.Open "SELECT PRD.PRD_IDY,SUM(tot_pcs),grp_cat,GRP_CLS,BRD_NME,inv.sub_unt FROM INV,prd,pgr WHERE inv.prd_idy like prd.prd_idy and prd.grp_idy like pgr.grp_idy and INV_IDY IN(SELECT INV_IDY FROM INL WHERE INV_DAT <= CTOD('" & DATE1 & "')  AND INV_DAT >=CTOD('" & MINDAT & "') AND TXN_TYP LIKE 'SAL' AND ROU_IDY LIKE '" & Trim(ROUIDY) & "')AND INV.STK_TYP IN('Salable','Free') GROUP BY INV.PRD_IDY", CON, adOpenKeyset, adLockOptimistic
End If
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
    If list1 = "Volume" Then
        CON.Execute "UPDATE trt SET a_mth_stk = " & Val(rs(1) & "") & " WHERE PRD_IDY LIKE '" & Trim(rs(0)) & "'", d
    Else
        CON.Execute "UPDATE trt SET a_mth_stk = " & STM(Val(rs(1) & ""), Val(rs(5) & "")) & " WHERE PRD_IDY LIKE '" & Trim(rs(0)) & "'", d
    End If
    If d = 0 Then
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT", CON, adOpenKeyset, adLockOptimistic
        TRT.AddNew
        TRT!prd_idy = Trim(rs(0))
        TRT!prd_cat = Trim(rs(2))
        TRT!prd_cls = Trim(rs(3))
        TRT!BRD_NME = Trim(rs(4))
        If Trim(list1) = "Volume" Then
            TRT!a_mth_stk = Val(rs(1) & "")
        Else
            TRT!a_mth_stk = STM(Val(rs(1) & ""), Val(rs(5) & ""))
        End If
        TRT!a_wek_stk = 0
        TRT!a_DAY_stk = 0
        TRT!T_DAY_STK = 0
        TRT!T_WEK_STK = 0
        TRT!T_MTH_STK = 0
        TRT!MTH_DYS = DateDiff("D", MINDAT, TODAT) + 1
        TRT!WEK_DYS = DAYS
        TRT!CMP_WEK = DateDiff("D", FRMDAT, DATE1) + 1
        TRT!CMP_MTH = DateDiff("D", MINDAT, DATE1) + 1
        TRT.Update
    End If
    rs.MoveNext
    Next
End If
If Me.Tag = "PNT" Then
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM ACT WHERE ACT_TYP IN('BOT','SAL') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "'", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET MTH_DYS = " & Val(rs.RecordCount), d
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TTL'", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs("DIR_IDY") = "SYS"
        rs("DIR_IDX") = "TTL"
        If list1 = "Quantity" Then
            rs(3) = "Retailer Penetration Report - Cases"
        Else
            rs(3) = "Retailer Penetration Report - Liters"
        End If
        rs(2) = ""
        rs.Update
    Else
        If list1 = "Quantity" Then
            CON.Execute "UPDATE DIR SET DIR_MSG = 'Retailer Penetration Report - Cases' WHERE DIR_IDX LIKE 'TTL'"
        Else
            CON.Execute "UPDATE DIR SET DIR_MSG = 'Retailer Penetration Report - Liters' WHERE DIR_IDX LIKE 'TTL'"
        End If
    End If
        If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'PND'", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs("DIR_IDY") = "SYS"
        rs("DIR_IDX") = "PND"
        rs(2) = ""
        rs(3) = ""
        rs.Update
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'PNW'", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs("DIR_IDY") = "SYS"
        rs("DIR_IDX") = "PNW"
        rs(2) = ""
        rs(3) = ""
        rs.Update
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'PNM'", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs("DIR_IDY") = "SYS"
        rs("DIR_IDX") = "PNM"
        rs(2) = ""
        rs(3) = ""
        rs.Update
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM INL WHERE INV_DAT = CTOD('" & DATE1 & "') AND TXN_TYP LIKE 'SAL' AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then CON.Execute "UPDATE DIR SET DIR_ALS = '" & Val(rs.RecordCount) & "' WHERE DIR_IDX LIKE 'PND'"
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM INL WHERE INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & FRMDAT & "') AND TXN_TYP LIKE 'SAL' AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then CON.Execute "UPDATE DIR SET DIR_ALS = '" & Val(rs.RecordCount) & "' WHERE DIR_IDX LIKE 'PNW'"
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM INL WHERE INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & MINDAT & "') AND TXN_TYP LIKE 'SAL' AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then CON.Execute "UPDATE DIR SET DIR_ALS = '" & Val(rs.RecordCount) & "' WHERE DIR_IDX LIKE 'PNM'"
    
    
    If Option1.Value = True Then           'Category
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT GROUP BY PRD_CAT", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD,PGR WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND INV_DAT = CTOD('" & DATE1 & "') AND TXN_TYP LIKE 'SAL' AND GRP_CAT LIKE '" & Trim(TRT!prd_cat) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_DAY_STK = " & Val(rs.RecordCount & "") & " WHERE PRD_CAT LIKE '" & Trim(TRT!prd_cat) & "'", d
            TRT.MoveNext
        Next
        
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT GROUP BY PRD_CAT", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD,PGR WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & FRMDAT & "') AND TXN_TYP LIKE 'SAL' AND GRP_CAT LIKE '" & Trim(TRT!prd_cat) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_WEK_STK = " & Val(rs.RecordCount & "") & " WHERE PRD_CAT LIKE '" & Trim(TRT!prd_cat) & "'", d
            TRT.MoveNext
        Next
        
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT GROUP BY PRD_CAT", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD,PGR WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & MINDAT & "') AND TXN_TYP LIKE 'SAL' AND GRP_CAT LIKE '" & Trim(TRT!prd_cat) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_MTH_STK = " & Val(rs.RecordCount & "") & " WHERE PRD_cat LIKE '" & Trim(TRT!prd_cat) & "'", d
            TRT.MoveNext
        Next
        TRT.Close
        rs.Close
        'Openreport "", CURDIR & "\Retailer Vol. Penetration.RRW", 0, "", 6
        LOADPRODUCTIVITY "Category"
        PRD_ANA_FRM.Tag = "Category"
        PRD_ANA_FRM.Show 1
        
    ElseIf Option2.Value = True Then       'Segment
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT GROUP BY PRD_CLS", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD,PGR WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND INV_DAT = CTOD('" & DATE1 & "') AND TXN_TYP LIKE 'SAL' AND GRP_CLS LIKE '" & Trim(TRT!prd_cls) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_DAY_STK = " & Val(rs.RecordCount & "") & " WHERE PRD_CLS LIKE '" & Trim(TRT!prd_cls) & "'", d
            TRT.MoveNext
        Next
        
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT group by prd_cls", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD,PGR WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & FRMDAT & "') AND TXN_TYP LIKE 'SAL' AND GRP_CLS LIKE '" & Trim(TRT!prd_cls) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_WEK_STK = " & Val(rs.RecordCount & "") & " WHERE PRD_CLS LIKE '" & Trim(TRT!prd_cls) & "'", d
            TRT.MoveNext
        Next
        
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT group by prd_cls", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD,PGR WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND PRD.GRP_IDY LIKE PGR.GRP_IDY AND INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & MINDAT & "') AND TXN_TYP LIKE 'SAL' AND grp_cls LIKE '" & Trim(TRT!prd_cls) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_MTH_STK = " & Val(rs.RecordCount & "") & " WHERE PRD_CLS LIKE '" & Trim(TRT!prd_cls) & "'", d
            TRT.MoveNext
        Next
        'Openreport CURDIR & "\cr.rp1", "Retailer Vol. Penetration Seg", 0, "", 6
        'Openreport "", CURDIR & "\Retailer Vol. Penetration Seg.RRW", 0, "", 6
        LOADPRODUCTIVITY "Segment"
        PRD_ANA_FRM.Tag = "Segment"
        PRD_ANA_FRM.Show 1
    ElseIf Option3.Value = True Then       'Brand
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT GROUP BY BRD_NME", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND INV_DAT = CTOD('" & DATE1 & "') AND TXN_TYP LIKE 'SAL' AND BRD_NME LIKE '" & Trim(TRT!BRD_NME) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_DAY_STK = " & Val(rs.RecordCount & "") & " WHERE BRD_NME LIKE '" & Trim(TRT!BRD_NME) & "'", d
            TRT.MoveNext
        Next
        
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT GROUP BY BRD_NME", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & FRMDAT & "') AND TXN_TYP LIKE 'SAL' AND BRD_NME LIKE '" & Trim(TRT!BRD_NME) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_WEK_STK = " & Val(rs.RecordCount & "") & " WHERE BRD_NME LIKE '" & Trim(TRT!BRD_NME) & "'", d
            TRT.MoveNext
        Next
        
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT GROUP BY BRD_NME", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV,PRD WHERE INV.PRD_IDY LIKE PRD.PRD_IDY AND INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & MINDAT & "') AND TXN_TYP LIKE 'SAL' AND BRD_NME LIKE '" & Trim(TRT!BRD_NME) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_MTH_STK = " & Val(rs.RecordCount & "") & " WHERE BRD_NME LIKE '" & Trim(TRT!BRD_NME) & "'", d
            TRT.MoveNext
        Next
        'Openreport CURDIR & "\cr.rp1", "Retailer Vol. Penetration Brd", 0, "", 6
        'Openreport "", CURDIR & "\Retailer Vol. Penetration Brd.RRW", 0, "", 6
        LOADPRODUCTIVITY "Brand"
        PRD_ANA_FRM.Tag = "Brand"
        PRD_ANA_FRM.Show 1
    Else
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV WHERE INV_DAT = CTOD('" & DATE1 & "') AND TXN_TYP LIKE 'SAL' AND PRD_IDY LIKE '" & Trim(TRT!prd_idy) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_DAY_STK = " & Val(rs.RecordCount & "") & " WHERE PRD_IDY LIKE '" & Trim(TRT!prd_idy) & "'", d
            TRT.MoveNext
        Next
        
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV WHERE INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & FRMDAT & "') AND TXN_TYP LIKE 'SAL' AND PRD_IDY LIKE '" & Trim(TRT!prd_idy) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_WEK_STK = " & Val(rs.RecordCount & "") & " WHERE PRD_IDY LIKE '" & Trim(TRT!prd_idy) & "'", d
            TRT.MoveNext
        Next
        
        If TRT.State = 1 Then TRT.Close
        TRT.Open "SELECT * FROM TRT", CON, adOpenKeyset
        For A = 1 To TRT.RecordCount
            If rs.State = 1 Then rs.Close
            rs.Open "SELECT * FROM INL WHERE INV_IDY IN(SELECT INV_IDY FROM INV WHERE INV_DAT <= CTOD('" & DATE1 & "') AND INV_DAT >=CTOD('" & MINDAT & "') AND TXN_TYP LIKE 'SAL' AND PRD_IDY LIKE '" & Trim(TRT!prd_idy) & "') AND ROU_IDY LIKE '" & Trim(ROUIDY) & "' GROUP BY CST_IDY", CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then CON.Execute "UPDATE TRT SET T_MTH_STK = " & Val(rs.RecordCount & "") & " WHERE PRD_IDY LIKE '" & Trim(TRT!prd_idy) & "'", d
            TRT.MoveNext
        Next
        'Openreport CURDIR & "\cr.rp1", "Retailer Vol. Penetration Prd", 0, "", 6
        'Openreport "", CURDIR & "\Retailer Vol. Penetration Prd.RRW", 0, "", 6
        LOADPRODUCTIVITY "Product"
        PRD_ANA_FRM.Tag = "Product"
        PRD_ANA_FRM.Show 1
    End If
Else
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM TGT", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then tgt = "Product"
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM TGS", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then tgt = "Segment"
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM TGB", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then tgt = "Brand"
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM TGC", CON, adOpenKeyset, adLockOptimistic
    If rs.RecordCount > 0 Then tgt = "Category"

    If Option1.Value = True Then    'CAT
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tgc where wek_idy = " & Val(wekidy), CON, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                If list1 = "Volume" Then
                    CON.Execute "UPDATE trt where prd_cat like '" & Trim(rs!pgt_cat) & "' SET t_wek_stk = " & Val(rs!TGT_CFC), d
                    CON.Execute "UPDATE trt where prd_cat like '" & Trim(rs!pgt_cat) & "' SET t_day_STK = " & Val(Val(rs!TGT_CFC) / DAYS), d
                Else
                    CON.Execute "UPDATE trt where prd_cat like '" & Trim(rs!pgt_cat) & "' SET t_wek_stk = " & Val(rs!TGT_QTY), d
                    CON.Execute "UPDATE trt where prd_cat like '" & Trim(rs!pgt_cat) & "' SET t_day_STK = " & Val(Val(rs!TGT_QTY) / DAYS), d
                End If
                rs.MoveNext
            Next
        End If
    
        If rs.State = 1 Then rs.Close
        If list1 = "Volume" Then
            rs.Open "select SUM(TGT_CFC),PGT_CAT from tgc where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by pgt_cat", CON, adOpenKeyset, adLockOptimistic
        Else
            rs.Open "select SUM(TGT_QTY),PGT_CAT from tgc where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by pgt_cat", CON, adOpenKeyset, adLockOptimistic
        End If
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                CON.Execute "UPDATE trt where ALLTRIM(prd_cat) like '" & Trim(rs!pgt_cat) & "' SET t_MTH_stk = " & Val(rs(0) & ""), d
                rs.MoveNext
            Next
        End If
        If tgt = "Segment" Then         'in Category Segment updation
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tgs where wek_idy = " & Val(wekidy), CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then
                For A = 1 To rs.RecordCount
                    If list1 = "Volume" Then
                        CON.Execute "UPDATE trt where prd_cls like '" & Trim(rs!pgr_seg) & "' SET t_wek_stk = " & Val(rs!TGT_CFC), d
                        CON.Execute "UPDATE trt where prd_cls like '" & Trim(rs!pgr_seg) & "' SET t_day_STK = " & Val(Val(rs!TGT_CFC) / DAYS), d
                    Else
                        CON.Execute "UPDATE trt where prd_cls like '" & Trim(rs!pgr_seg) & "' SET t_wek_stk = " & Val(rs!TGT_QTY), d
                        CON.Execute "UPDATE trt where prd_cls like '" & Trim(rs!pgr_seg) & "' SET t_day_STK = " & Val(Val(rs!TGT_QTY) / DAYS), d
                    End If
                    rs.MoveNext
                Next
            End If
        
            If rs.State = 1 Then rs.Close
            If list1 = "Volume" Then
                rs.Open "select SUM(TGT_CFC),pgr_seg from tgs where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by pgr_seg", CON, adOpenKeyset, adLockOptimistic
            Else
                rs.Open "select SUM(TGT_QTY),pgr_seg from tgs where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by pgr_seg", CON, adOpenKeyset, adLockOptimistic
            End If
            If rs.RecordCount > 0 Then
                For A = 1 To rs.RecordCount
                    CON.Execute "UPDATE trt where prd_cls like '" & Trim(rs!pgr_seg) & "' SET t_MTH_stk = " & Val(rs(0) & ""), d
                    rs.MoveNext
                Next
            End If
            
            If rs.State = 1 Then rs.Close
            rs.Open "select * from trt group by prd_cls", CON, adOpenKeyset, adLockOptimistic
            ReDim ary(0)
            ReDim aryday(0)
            ReDim arywek(0)
            ReDim arymth(0)
            For A = 0 To rs.RecordCount - 1
                t = UBound(ary)
                For B = 0 To t
                If Trim(ary(B)) = Trim(rs!prd_cat) Then
                    aryday(B) = Val(rs!T_DAY_STK) + aryday(B)
                    arywek(B) = Val(rs!T_WEK_STK) + arywek(B)
                    arymth(B) = Val(rs!T_MTH_STK) + arymth(B)
                    GoTo nxt
                End If
                Next
                ReDim Preserve ary(t + 1)
                ReDim Preserve aryday(t + 1)
                ReDim Preserve arywek(t + 1)
                ReDim Preserve arymth(t + 1)
                ary(t) = Trim(rs!prd_cat)
                aryday(t) = Val(rs!T_DAY_STK)
                arywek(t) = Val(rs!T_WEK_STK)
                arymth(t) = Val(rs!T_MTH_STK)
nxt:
                rs.MoveNext
            Next
            For A = 0 To UBound(ary)
                CON.Execute "UPDATE trt where prd_cat like '" & Trim(ary(A)) & "' SET t_day_stk = " & Val(aryday(A) & ""), d
                CON.Execute "UPDATE trt where prd_cat like '" & Trim(ary(A)) & "' SET t_wek_stk = " & Val(arywek(A) & ""), d
                CON.Execute "UPDATE trt where prd_cat like '" & Trim(ary(A)) & "' SET t_mth_stk = " & Val(arymth(A) & ""), d
            Next
        ElseIf tgt = "Brand" Then 'in Category  Brand updation
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tgb where wek_idy = " & Val(wekidy), CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then
                For A = 1 To rs.RecordCount
                    If list1 = "Volume" Then
                        CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_wek_stk = " & Val(rs!TGT_CFC), d
                        CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_day_STK = " & Val(Val(rs!TGT_CFC) / DAYS), d
                    Else
                        CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_wek_stk = " & Val(rs!TGT_QTY), d
                        CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_day_STK = " & Val(Val(rs!TGT_QTY) / DAYS), d
                    End If
                    rs.MoveNext
                Next
            End If
        
            If rs.State = 1 Then rs.Close
            If list1 = "Volume" Then
                rs.Open "select SUM(TGT_CFC),BRD_NME from tgb where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by brd_nme", CON, adOpenKeyset, adLockOptimistic
            Else
                rs.Open "select SUM(TGT_QTY),BRD_NME from tgb where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by brd_nme", CON, adOpenKeyset, adLockOptimistic
            End If
            If rs.RecordCount > 0 Then
                For A = 1 To rs.RecordCount
                    CON.Execute "UPDATE trt where ALLTRIM(brd_nme) like '" & Trim(rs!BRD_NME) & "' SET t_MTH_stk = " & Val(rs(0) & ""), d
                    rs.MoveNext
                Next
            End If
            
            If rs.State = 1 Then rs.Close
            rs.Open "select * from trt group by brd_nme", CON, adOpenKeyset, adLockOptimistic
            ReDim ary(0)
            ReDim aryday(0)
            ReDim arywek(0)
            ReDim arymth(0)
            For A = 0 To rs.RecordCount - 1
                t = UBound(ary)
                For B = 0 To t
                If Trim(ary(B)) = Trim(rs!prd_cat) Then
                    aryday(B) = Val(rs!T_DAY_STK) + aryday(B)
                    arywek(B) = Val(rs!T_WEK_STK) + arywek(B)
                    arymth(B) = Val(rs!T_MTH_STK) + arymth(B)
                    GoTo nxt1
                End If
                Next
                ReDim Preserve ary(t + 1)
                ReDim Preserve aryday(t + 1)
                ReDim Preserve arywek(t + 1)
                ReDim Preserve arymth(t + 1)
                ary(t) = Trim(rs!prd_cat)
                aryday(t) = Val(rs!T_DAY_STK)
                arywek(t) = Val(rs!T_WEK_STK)
                arymth(t) = Val(rs!T_MTH_STK)
nxt1:
                rs.MoveNext
            Next
            For A = 0 To UBound(ary)
                CON.Execute "UPDATE trt where prd_cat like '" & Trim(ary(A)) & "' SET t_day_stk = " & Val(aryday(A) & ""), d
                CON.Execute "UPDATE trt where prd_cat like '" & Trim(ary(A)) & "' SET t_wek_stk = " & Val(arywek(A) & ""), d
                CON.Execute "UPDATE trt where prd_cat like '" & Trim(ary(A)) & "' SET t_mth_stk = " & Val(arymth(A) & ""), d
            Next
        ElseIf tgt = "Product" Then 'in Category  Product updation
            If rs.State = 1 Then rs.Close
        rs.Open "select * from tgt where wek_idy = " & Val(wekidy), CON, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                If list1 = "Volume" Then
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_wek_stk = " & Val(rs!TGT_CFC), d
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_day_STK = " & Val(Val(rs!TGT_CFC) / DAYS), d
                Else
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_wek_stk = " & Val(rs!TGT_QTY), d
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_day_STK = " & Val(Val(rs!TGT_QTY) / DAYS), d
                End If
                rs.MoveNext
            Next
        End If
    
        If rs.State = 1 Then rs.Close
        If list1 = "Volume" Then
            rs.Open "select SUM(TGT_CFC),prd_idy from tgt where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by prd_idy", CON, adOpenKeyset, adLockOptimistic
        Else
            rs.Open "select SUM(TGT_QTY),prd_idy from tgt where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by prd_idy", CON, adOpenKeyset, adLockOptimistic
        End If
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_MTH_stk = " & Val(rs(0) & ""), d
                rs.MoveNext
            Next
        End If
            
            If rs.State = 1 Then rs.Close
            rs.Open "select * from trt group by prd_idy", CON, adOpenKeyset, adLockOptimistic
            ReDim ary(0)
            ReDim aryday(0)
            ReDim arywek(0)
            ReDim arymth(0)
            For A = 0 To rs.RecordCount - 1
                MsgBox rs!prd_idy
                t = UBound(ary)
                For B = 0 To t
                If Trim(ary(B)) = Trim(rs!prd_cat) Then
                    aryday(B) = Val(rs!T_DAY_STK) + aryday(B)
                    arywek(B) = Val(rs!T_WEK_STK) + arywek(B)
                    arymth(B) = Val(rs!T_MTH_STK) + arymth(B)
                    GoTo nxt2
                End If
                Next
                ReDim Preserve ary(t + 1)
                ReDim Preserve aryday(t + 1)
                ReDim Preserve arywek(t + 1)
                ReDim Preserve arymth(t + 1)
                ary(t) = Trim(rs!prd_cat)
                aryday(t) = Val(rs!T_DAY_STK)
                arywek(t) = Val(rs!T_WEK_STK)
                arymth(t) = Val(rs!T_MTH_STK)
nxt2:
                rs.MoveNext
            Next
            For A = 0 To UBound(ary)
                CON.Execute "UPDATE trt where prd_cat like '" & Trim(ary(A)) & "' SET t_day_stk = " & Val(aryday(A) & ""), d
                CON.Execute "UPDATE trt where prd_cat like '" & Trim(ary(A)) & "' SET t_wek_stk = " & Val(arywek(A) & ""), d
                CON.Execute "UPDATE trt where prd_cat like '" & Trim(ary(A)) & "' SET t_mth_stk = " & Val(arymth(A) & ""), d
            Next
        End If
        Openreport CURDIR & "\cr.rp1", "Category Targets", 0, "", 6
    ElseIf Option2.Value = True Then    'SEG
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tgs where wek_idy = " & Val(wekidy), CON, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                If list1 = "Volume" Then
                    CON.Execute "UPDATE trt where prd_cls like '" & Trim(rs!pgr_seg) & "' SET t_wek_stk = " & Val(rs!TGT_CFC), d
                    CON.Execute "UPDATE trt where prd_cls like '" & Trim(rs!pgr_seg) & "' SET t_day_STK = " & Val(Val(rs!TGT_CFC) / DAYS), d
                Else
                    CON.Execute "UPDATE trt where prd_cls like '" & Trim(rs!pgr_seg) & "' SET t_wek_stk = " & Val(rs!TGT_QTY), d
                    CON.Execute "UPDATE trt where prd_cls like '" & Trim(rs!pgr_seg) & "' SET t_day_STK = " & Val(Val(rs!TGT_QTY) / DAYS), d
                End If
                rs.MoveNext
            Next
        End If
        If rs.State = 1 Then rs.Close
        If list1 = "Volume" Then
            rs.Open "select SUM(TGT_CFC),pgr_seg from tgs where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by pgr_seg", CON, adOpenKeyset, adLockOptimistic
        Else
            rs.Open "select SUM(TGT_QTY),pgr_seg from tgs where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by pgr_seg", CON, adOpenKeyset, adLockOptimistic
        End If
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                CON.Execute "UPDATE trt where prd_cls like '" & Trim(rs!pgr_seg) & "' SET t_MTH_stk = " & Val(rs(0) & ""), d
                rs.MoveNext
            Next
        End If
        
        If tgt = "Brand" Then 'in Segment Brand updation
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tgb where wek_idy = " & Val(wekidy), CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then
                For A = 1 To rs.RecordCount
                    If list1 = "Volume" Then
                        CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_wek_stk = " & Val(rs!TGT_CFC), d
                        CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_day_STK = " & Val(Val(rs!TGT_CFC) / DAYS), d
                    Else
                        CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_wek_stk = " & Val(rs!TGT_QTY), d
                        CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_day_STK = " & Val(Val(rs!TGT_QTY) / DAYS), d
                    End If
                    rs.MoveNext
                Next
            End If
        
            If rs.State = 1 Then rs.Close
            If list1 = "Volume" Then
                rs.Open "select SUM(TGT_CFC),BRD_NME from tgb where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by brd_nme", CON, adOpenKeyset, adLockOptimistic
            Else
                rs.Open "select SUM(TGT_QTY),BRD_NME from tgb where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by brd_nme", CON, adOpenKeyset, adLockOptimistic
            End If
            If rs.RecordCount > 0 Then
                For A = 1 To rs.RecordCount
                    CON.Execute "UPDATE trt where ALLTRIM(brd_nme) like '" & Trim(rs!BRD_NME) & "' SET t_MTH_stk = " & Val(rs(0) & ""), d
                    rs.MoveNext
                Next
            End If
            
            If rs.State = 1 Then rs.Close
            rs.Open "select * from trt group by brd_nme", CON, adOpenKeyset, adLockOptimistic
            ReDim ary(0)
            ReDim aryday(0)
            ReDim arywek(0)
            ReDim arymth(0)
            For A = 0 To rs.RecordCount - 1
                t = UBound(ary)
                For B = 0 To t
                If Trim(ary(B)) = Trim(rs!prd_cls) Then
                    aryday(B) = Val(rs!T_DAY_STK) + aryday(B)
                    arywek(B) = Val(rs!T_WEK_STK) + arywek(B)
                    arymth(B) = Val(rs!T_MTH_STK) + arymth(B)
                    GoTo nxt3
                End If
                Next
                ReDim Preserve ary(t + 1)
                ReDim Preserve aryday(t + 1)
                ReDim Preserve arywek(t + 1)
                ReDim Preserve arymth(t + 1)
                ary(t) = Trim(rs!prd_cls)
                aryday(t) = Val(rs!T_DAY_STK)
                arywek(t) = Val(rs!T_WEK_STK)
                arymth(t) = Val(rs!T_MTH_STK)
nxt3:
                rs.MoveNext
            Next

            For A = 0 To UBound(ary)
                CON.Execute "UPDATE trt where prd_cls like '" & Trim(ary(A)) & "' SET t_day_stk = " & Val(aryday(A) & ""), d
                CON.Execute "UPDATE trt where prd_cls like '" & Trim(ary(A)) & "' SET t_wek_stk = " & Val(arywek(A) & ""), d
                CON.Execute "UPDATE trt where prd_cls like '" & Trim(ary(A)) & "' SET t_mth_stk = " & Val(arymth(A) & ""), d
            Next
        ElseIf tgt = "Product" Then 'in Segment  Product updation
            If rs.State = 1 Then rs.Close
        rs.Open "select * from tgt where wek_idy = " & Val(wekidy), CON, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                If list1 = "Volume" Then
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_wek_stk = " & Val(rs!TGT_CFC), d
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_day_STK = " & Val(Val(rs!TGT_CFC) / DAYS), d
                Else
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_wek_stk = " & Val(rs!TGT_QTY), d
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_day_STK = " & Val(Val(rs!TGT_QTY) / DAYS), d
                End If
                rs.MoveNext
            Next
        End If
    
        If rs.State = 1 Then rs.Close
        If list1 = "Volume" Then
            rs.Open "select SUM(TGT_CFC),prd_idy from tgt where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by prd_idy", CON, adOpenKeyset, adLockOptimistic
        Else
            rs.Open "select SUM(TGT_QTY),prd_idy from tgt where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by prd_idy", CON, adOpenKeyset, adLockOptimistic
        End If
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_MTH_stk = " & Val(rs(0) & ""), d
                rs.MoveNext
            Next
        End If
            
            If rs.State = 1 Then rs.Close
            rs.Open "select * from trt group by prd_idy", CON, adOpenKeyset, adLockOptimistic
            ReDim ary(0)
            ReDim aryday(0)
            ReDim arywek(0)
            ReDim arymth(0)
            For A = 0 To rs.RecordCount - 1
                t = UBound(ary)
                For B = 0 To t
                If Trim(ary(B)) = Trim(rs!prd_cls) Then
                    aryday(B) = Val(rs!T_DAY_STK) + aryday(B)
                    arywek(B) = Val(rs!T_WEK_STK) + arywek(B)
                    arymth(B) = Val(rs!T_MTH_STK) + arymth(B)
                    GoTo nxt4
                End If
                Next
                ReDim Preserve ary(t + 1)
                ReDim Preserve aryday(t + 1)
                ReDim Preserve arywek(t + 1)
                ReDim Preserve arymth(t + 1)
                ary(t) = Trim(rs!prd_cls)
                aryday(t) = Val(rs!T_DAY_STK)
                arywek(t) = Val(rs!T_WEK_STK)
                arymth(t) = Val(rs!T_MTH_STK)
nxt4:
                rs.MoveNext
            Next
            For A = 0 To UBound(ary)
                CON.Execute "UPDATE trt where prd_cls like '" & Trim(ary(A)) & "' SET t_day_stk = " & Val(aryday(A) & ""), d
                CON.Execute "UPDATE trt where prd_cls like '" & Trim(ary(A)) & "' SET t_wek_stk = " & Val(arywek(A) & ""), d
                CON.Execute "UPDATE trt where prd_cls like '" & Trim(ary(A)) & "' SET t_mth_stk = " & Val(arymth(A) & ""), d
            Next
        End If
        Openreport CURDIR & "\cr.rp1", "Segment Targets", 0, "", 6
    ElseIf Option3.Value = True Then    'BRAND
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tgb where wek_idy = " & Val(wekidy), CON, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                If list1 = "Volume" Then
                    CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_wek_stk = " & Val(rs!TGT_CFC), d
                    CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_day_STK = " & Val(Val(rs!TGT_CFC) / DAYS), d
                Else
                    CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_wek_stk = " & Val(rs!TGT_QTY), d
                    CON.Execute "UPDATE trt where BRD_NME like '" & Trim(rs!BRD_NME) & "' SET t_day_STK = " & Val(Val(rs!TGT_QTY) / DAYS), d
                End If
                rs.MoveNext
            Next
        End If
    
        If rs.State = 1 Then rs.Close
        If list1 = "Volume" Then
            rs.Open "select SUM(TGT_CFC),BRD_NME from tgb where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by brd_nme", CON, adOpenKeyset, adLockOptimistic
        Else
            rs.Open "select SUM(TGT_QTY),BRD_NME from tgb where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by brd_nme", CON, adOpenKeyset, adLockOptimistic
        End If
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                CON.Execute "UPDATE trt where ALLTRIM(brd_nme) like '" & Trim(rs!BRD_NME) & "' SET t_MTH_stk = " & Val(rs(0) & ""), d
                rs.MoveNext
            Next
        End If
        
        If tgt = "Product" Then 'in Brand Product updation
            If rs.State = 1 Then rs.Close
            rs.Open "select * from tgt where wek_idy = " & Val(wekidy), CON, adOpenKeyset, adLockOptimistic
            If rs.RecordCount > 0 Then
                For A = 1 To rs.RecordCount
                    If list1 = "Volume" Then
                        CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_wek_stk = " & Val(rs!TGT_CFC), d
                        CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_day_STK = " & Val(Val(rs!TGT_CFC) / DAYS), d
                    Else
                        CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_wek_stk = " & Val(rs!TGT_QTY), d
                        CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_day_STK = " & Val(Val(rs!TGT_QTY) / DAYS), d
                    End If
                    rs.MoveNext
                Next
            End If
        
            If rs.State = 1 Then rs.Close
            If list1 = "Volume" Then
                rs.Open "select SUM(TGT_CFC),prd_idy from tgt where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by prd_idy", CON, adOpenKeyset, adLockOptimistic
            Else
                rs.Open "select SUM(TGT_QTY),prd_idy from tgt where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by prd_idy", CON, adOpenKeyset, adLockOptimistic
            End If
            If rs.RecordCount > 0 Then
                For A = 1 To rs.RecordCount
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_MTH_stk = " & Val(rs(0) & ""), d
                    rs.MoveNext
                Next
            End If
                
                If rs.State = 1 Then rs.Close
                rs.Open "select * from trt group by prd_idy", CON, adOpenKeyset, adLockOptimistic
                ReDim ary(0)
                ReDim aryday(0)
                ReDim arywek(0)
                ReDim arymth(0)
                For A = 0 To rs.RecordCount - 1
                    t = UBound(ary)
                    For B = 0 To t
                    If Trim(ary(B)) = Trim(rs!BRD_NME) Then
                        aryday(B) = Val(rs!T_DAY_STK) + aryday(B)
                        arywek(B) = Val(rs!T_WEK_STK) + arywek(B)
                        arymth(B) = Val(rs!T_MTH_STK) + arymth(B)
                        GoTo nxt5
                    End If
                    Next
                    ReDim Preserve ary(t + 1)
                    ReDim Preserve aryday(t + 1)
                    ReDim Preserve arywek(t + 1)
                    ReDim Preserve arymth(t + 1)
                    ary(t) = Trim(rs!BRD_NME)
                    aryday(t) = Val(rs!T_DAY_STK)
                    arywek(t) = Val(rs!T_WEK_STK)
                    arymth(t) = Val(rs!T_MTH_STK)
nxt5:
                    rs.MoveNext
                Next
                For A = 0 To UBound(ary)
                    CON.Execute "UPDATE trt where BRD_NME like '" & Trim(ary(A)) & "' SET t_day_stk = " & Val(aryday(A) & ""), d
                    CON.Execute "UPDATE trt where BRD_NME like '" & Trim(ary(A)) & "' SET t_wek_stk = " & Val(arywek(A) & ""), d
                    CON.Execute "UPDATE trt where BRD_NME like '" & Trim(ary(A)) & "' SET t_mth_stk = " & Val(arymth(A) & ""), d
                Next
            End If
        Openreport CURDIR & "\cr.rp1", "Brand Targets", 0, "", 6
    Else    'SKU
        If rs.State = 1 Then rs.Close
        rs.Open "select * from tgt where wek_idy = " & Val(wekidy), CON, adOpenKeyset, adLockOptimistic
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                If list1 = "Volume" Then
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_wek_stk = " & Val(rs!TGT_CFC), d
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_day_STK = " & Val(Val(rs!TGT_CFC) / DAYS), d
                Else
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_wek_stk = " & Val(rs!TGT_QTY), d
                    CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_day_STK = " & Val(Val(rs!TGT_QTY) / DAYS), d
                End If
                rs.MoveNext
            Next
        End If
    
        If rs.State = 1 Then rs.Close
        If list1 = "Volume" Then
            rs.Open "select SUM(TGT_CFC),prd_idy from tgt where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by prd_idy", CON, adOpenKeyset, adLockOptimistic
        Else
            rs.Open "select SUM(TGT_QTY),prd_idy from tgt where wek_idy IN(SELECT WEK_IDY FROM WEK WHERE MTH_NOS = " & Val(MTHIDY) & ")group by prd_idy", CON, adOpenKeyset, adLockOptimistic
        End If
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
                CON.Execute "UPDATE trt where prd_idy like '" & Trim(rs!prd_idy) & "' SET t_MTH_stk = " & Val(rs(0) & ""), d
                rs.MoveNext
            Next
        End If
        Openreport CURDIR & "\cr.rp1", "Product Targets", 0, "", 6
    End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    ROU_LST.Selected(0) = True
    ROU_LST.SetFocus
End If
End Sub

Private Sub Form_Activate()
DATE1.SetFocus
ROU_LST.Selected(0) = True
list1.Selected(0) = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_load()
DATE1 = Date
End Sub

Private Sub List1_keypress(K As Integer)
If K = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub ROU_LST_kEYPRESS(K As Integer)
If K = 13 Then
    list1.SetFocus
End If
End Sub

Public Sub LOADPRODUCTIVITY(STR As String)
On Error Resume Next
Dim rs As New ADODB.Recordset
With PRD_ANA_FRM.MS
.ColWidth(0) = 2000
If STR = "Category" Then
    STR1 = "PRD_CAT"
    .TextMatrix(0, 0) = "Category"
ElseIf STR = "Segment" Then
    STR1 = "PRD_CLS"
    .TextMatrix(0, 0) = "Segment"
ElseIf STR = "Brand" Then
    STR1 = "BRD_NME"
    .TextMatrix(0, 0) = "Brand Name"
ElseIf STR = "Product" Then
    .TextMatrix(0, 0) = "Prdouct Name"
End If
.TextMatrix(0, 1) = "Day.Vol"
.TextMatrix(0, 2) = "Wek.Vol"
.TextMatrix(0, 3) = "Mth.Vol"
.TextMatrix(0, 4) = "Tot.Out"
.TextMatrix(0, 5) = "Prod.Day"
.TextMatrix(0, 6) = "Prod.Wek"
.TextMatrix(0, 7) = "Prod.Mth"
.TextMatrix(0, 8) = "% Pntr.Day"
.TextMatrix(0, 9) = "% Pntr.Wek"
.TextMatrix(0, 10) = "% Pntr.Mth"
.TextMatrix(0, 11) = "VPO.Day"
.TextMatrix(0, 12) = "VPO.Wek"
.TextMatrix(0, 13) = "VPO.Mth"
.TextMatrix(0, 14) = "% Cntr.MTD"
If STR = "Product" Then
    rs.Open "SELECT Prd_nme,SUM(A_DAY_STK) A_DAY_STK ,SUM(A_WEK_STK) A_WEK_STK,SUM(A_MTH_STK) A_MTH_STK,MTH_DYS,T_DAY_STK,T_WEK_STK,T_MTH_STK,ROUND((T_DAY_STK/MTH_DYS)*100,2) PNT_DAY,ROUND((T_WEK_STK/MTH_DYS)*100,2) PNT_WEK,ROUND((T_MTH_STK/MTH_DYS)*100,2) PNT_MTH FROM prd,TRT where trt.prd_idy like prd.prd_idy GROUP BY trt.prd_idy", CON, adOpenKeyset, adLockOptimistic
Else
    rs.Open "SELECT " & STR1 & ",SUM(A_DAY_STK) A_DAY_STK ,SUM(A_WEK_STK) A_WEK_STK,SUM(A_MTH_STK) A_MTH_STK,MTH_DYS,T_DAY_STK,T_WEK_STK,T_MTH_STK,ROUND((T_DAY_STK/MTH_DYS)*100,2) PNT_DAY,ROUND((T_WEK_STK/MTH_DYS)*100,2) PNT_WEK,ROUND((T_MTH_STK/MTH_DYS)*100,2) PNT_MTH FROM TRT GROUP BY " & STR1, CON, adOpenKeyset, adLockOptimistic
End If
.Rows = rs.RecordCount + 1
For A = 1 To rs.RecordCount
    For B = 0 To rs.Fields.Count - 1
        If B = 1 Or B = 2 Or B = 3 Then
            .TextMatrix(A, B) = QTY(Val(rs(B)))
        ElseIf B = 0 Then
            .TextMatrix(A, B) = Trim(rs(B))
        ElseIf B = 8 Or B = 9 Or B = 10 Then
            .TextMatrix(A, B) = convert(Val(rs(B)))
        Else
            .TextMatrix(A, B) = Val(rs(B))
        End If
    Next
    rs.MoveNext
Next
For A = 1 To .Rows - 1
    TOT_DAY = Val(.TextMatrix(A, 1)) + Val(TOT_DAY)
    TOT_WEK = Val(.TextMatrix(A, 2)) + Val(TOT_WEK)
    TOT_MTH = Val(.TextMatrix(A, 3)) + Val(TOT_MTH)
Next

For A = 1 To .Rows - 1
    If Val(.TextMatrix(A, 5)) = 0 Or Val(.TextMatrix(A, 1)) = 0 Then
        .TextMatrix(A, 11) = ""
    Else
        .TextMatrix(A, 11) = QTY(Round(Val(.TextMatrix(A, 1)) / Val(.TextMatrix(A, 5)), 3))
    End If
    If Val(.TextMatrix(A, 6)) = 0 Or Val(.TextMatrix(A, 2)) = 0 Then
        .TextMatrix(A, 12) = ""
    Else
        .TextMatrix(A, 12) = QTY(Round(Val(.TextMatrix(A, 2)) / Val(.TextMatrix(A, 6)), 3))
    End If
    If Val(.TextMatrix(A, 7)) = 0 Or Val(.TextMatrix(A, 3)) = 0 Then
        .TextMatrix(A, 13) = ""
    Else
        .TextMatrix(A, 13) = QTY(Round(Val(.TextMatrix(A, 3)) / Val(.TextMatrix(A, 7)), 3))
        .TextMatrix(A, 14) = discount(Round((Val(.TextMatrix(A, 3)) / Val(TOT_MTH)) * 100, 2))
    End If
Next
.AddItem ""
.Row = .Rows - 1
.Col = 0
.CellFontBold = True
.Col = 1
.CellFontBold = True
.Col = 2
.CellFontBold = True
.Col = 3
.CellFontBold = True
.Col = 4
.CellFontBold = True
.Col = 5
.CellFontBold = True
.Col = 6
.CellFontBold = True
.Col = 7
.CellFontBold = True
.Col = 8
.CellFontBold = True
.Col = 8
.CellFontBold = True
.Col = 9
.CellFontBold = True
.Col = 10
.CellFontBold = True

.TextMatrix(.Rows - 1, 0) = "Total"
.TextMatrix(.Rows - 1, 1) = QTY(Val(TOT_DAY))
.TextMatrix(.Rows - 1, 2) = QTY(Val(TOT_WEK))
.TextMatrix(.Rows - 1, 3) = QTY(Val(TOT_MTH))
.TextMatrix(.Rows - 1, 4) = .TextMatrix(1, 4)
.TextMatrix(.Rows - 1, 5) = GetValue("DIR_ALS", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'PND'")
.TextMatrix(.Rows - 1, 6) = GetValue("DIR_ALS", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'PNW'")
.TextMatrix(.Rows - 1, 7) = GetValue("DIR_ALS", "DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'PNM'")
.TextMatrix(.Rows - 1, 8) = Round((Val(.TextMatrix(.Rows - 1, 5)) / Val(.TextMatrix(.Rows - 1, 4))) * 100, 2)
.TextMatrix(.Rows - 1, 9) = Round((Val(.TextMatrix(.Rows - 1, 6)) / Val(.TextMatrix(.Rows - 1, 4))) * 100, 2)
.TextMatrix(.Rows - 1, 10) = Round((Val(.TextMatrix(.Rows - 1, 7)) / Val(.TextMatrix(.Rows - 1, 4))) * 100, 2)

End With
End Sub
