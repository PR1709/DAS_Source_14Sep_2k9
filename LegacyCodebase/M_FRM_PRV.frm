VERSION 5.00
Begin VB.Form M_FRM_PRV 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Voucher"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   ControlBox      =   0   'False
   Icon            =   "M_FRM_PRV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      Caption         =   "Check to print voucher after record save"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "M_FRM_PRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim REC_SET As New ADODB.Recordset
If M_FRM_PRV.Tag = "LSL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRE_LSL'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRE_LSL"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "AST_INW" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_AST_INW'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_AST_INW"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "AST_OUT" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_AST_OUT'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_AST_OUT"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "ADJ" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_ADJ'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_ADJ"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "COV_FRM" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_COV'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_COV"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "IRT" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_IRT'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_IRT"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "ORD" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_ORD'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_ORD"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "SLP" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_SAL_CHG'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_SAL_CHG"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "SLR" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_SAL_RAT'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_SAL_RAT"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "OPS" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_OPS'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_OPS"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "DMI" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_DMI'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_DMI"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "DMR" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_DMR'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_DMR"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "PUR" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_PUR'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_PUR"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "RSI" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_RSI1'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_RSI1"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "PRT" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_PRT'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_PRT"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "SRT" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_SRT'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_SRT"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "PMT" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_PMT'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_PMT"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "RCT" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_RCT'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_RCT"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "OPN" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_OPN'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_OPN"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "JRN" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_JRN'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_JRN"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "CRN" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_CRN'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_CRN"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "DBN" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_DBN'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_DBN"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "SAL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_SAL'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_SAL"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "MAN" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'm_frm_man'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "m_frm_man"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "RSL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'm_van_sal'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "m_van_sal"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "MPR" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'm_frm_mpr'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "m_frm_mpr"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "RSR" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_RSR'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_RSR"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "STI" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_STI'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_STI"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "STO" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_STO'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_STO"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "PCS" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_PCS'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_PCS"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "PCR" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_PCS_PCS'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_PCS_PCS"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "VAN" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_VAN_SAL'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_VAN_SAL"
        REC_SET.Update
    End If
    Unload Me
ElseIf M_FRM_PRV.Tag = "PSH_SAL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_SAL_PSH'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_SAL_PSH"
        REC_SET.Update
    End If
    Unload Me
ElseIf Me.Tag = "SAL_SHL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'SAL_SHL'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "SAL_SHL"
        REC_SET.Update
    End If
    Unload Me
ElseIf Me.Tag = "PUR_SHL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'PUR_SHL'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "PUR_SHL"
        REC_SET.Update
    End If
    Unload Me
ElseIf Me.Tag = "SRN_SHL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'SRN_SHL'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "SRN_SHL"
        REC_SET.Update
    End If
    Unload Me
ElseIf Me.Tag = "PRT_SHL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'PRT_SHL'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "PRT_SHL"
        REC_SET.Update
    End If
    Unload Me
ElseIf Me.Tag = "STI_SHL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'STI_SHL'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "STI_SHL"
        REC_SET.Update
    End If
    Unload Me
ElseIf Me.Tag = "STO_SHL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'STO_SHL'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "STO_SHL"
        REC_SET.Update
    End If
    Unload Me
ElseIf Me.Tag = "TRS_SHL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'TRS_SHL'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "TRS_SHL"
        REC_SET.Update
    End If
    Unload Me
ElseIf Me.Tag = "SIN_SHL" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'SIN_SHL'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "SIN_SHL"
        REC_SET.Update
    End If
    Unload Me
ElseIf Me.Tag = "TRS" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_TRS'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_TRS"
        REC_SET.Update
    End If
    Unload Me
ElseIf Me.Tag = "SOR" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_SOR'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_SOR"
        REC_SET.Update
    End If
    Unload Me
ElseIf Me.Tag = "POR" Then
    If REC_SET.State = 1 Then REC_SET.Close
    If Check3.value = 0 Then
        CON.Execute "delete from prt where prt_nme like 'M_FRM_POR'"
    Else
        If REC_SET.State = 1 Then REC_SET.Close
        REC_SET.Open "select * from prt", CON, adOpenKeyset, adLockPessimistic
        REC_SET.AddNew
        REC_SET(0) = "M_FRM_POR"
        REC_SET.Update
    End If
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

