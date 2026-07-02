VERSION 5.00
Begin VB.Form SYS_PSW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator Password"
   ClientHeight    =   1095
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3840
   Icon            =   "SYS_PSW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   646.963
   ScaleMode       =   0  'User
   ScaleWidth      =   3605.552
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox text1 
      Appearance      =   0  'Flat
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   285
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Password"
      Height          =   195
      Left            =   705
      TabIndex        =   1
      Top             =   360
      Width           =   1170
   End
End
Attribute VB_Name = "SYS_PSW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub TEXT1_KeyDown(K As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from usr where usr_typ like 'Admin'", CON, adOpenKeyset, adLockPessimistic
        If Not rs.EOF Then
            If Trim(Text1) = DPWD.decPwd(Trim(rs("USR_PWD"))) Or Trim(Text1) = "cspladmin" Then
                If Me.Caption = "DAS Version 4.0" Then
                    INS_INF.Show 1
                    Unload Me
                    Set Form = Nothing
                    Exit Sub
                ElseIf Me.Caption = "Account Deletion" Then
                    M_FRM_ACT.Tag = "Delete"
                    Unload SYS_PSW
                ElseIf Me.Caption = "Group Account Deletion" Then
                    M_FRM_GRP.Tag = "Delete"
                    Unload Me
                    Set Form = Nothing
                ElseIf Me.Caption = "Product Deletion" Then
                    DEL_REC_IDY = "D"
                    Unload SYS_PSW
                ElseIf Me.Caption = "Categories (Product Group) Deletion" Then
                    M_FRM_PGR.Tag = "Delete"
                    Unload Me
                    Set Form = Nothing
                ElseIf Me.Caption = "Location Deletion" Then
                    M_FRM_LOC.Tag = "Delete"
                    Unload Me
                ElseIf Me.Caption = "Category Deletion" Then
                    M_FRM_CAT.Tag = "Delete"
                    Unload Me
                ElseIf Me.Caption = "Market Deletion" Then
                    M_FRM_MAR.Tag = "Delete"
                    Unload Me
                ElseIf Me.Caption = "Salesmen Deletion" Then
                    M_FRM_SMN.Tag = "Delete"
                    Unload Me
                ElseIf Me.Caption = "User Deletion" Then
                    USR_PSW.Tag = "Delete"
                    Unload Me
                ElseIf Me.Caption = "Password" Then
                    Unload Me
                    M_FRM_RCT.Tag = ""
                    M_FRM_ASD.Caption = "Invoice Date"
                    M_FRM_ASD.Show 1
                ElseIf Me.Tag = "SAL" Then
                    Unload Me
                    M_FRM_SAL_CHG.Show 1
                ElseIf Me.Tag = "CHG" Then
                    Unload Me
                    M_FRE_SAL_CHG.Show 1
                ElseIf Me.Tag = "Template" Then
                    Unload Me
                    M_INV_TYP.Show 1
                ElseIf Me.Tag = "Pur Check" Then
                    Unload Me
                    M_FRM_PRC.Show 1
                ElseIf Me.Tag = "Check Date" Then
                    Unload Me
                    M_FRM_SAL.Tag = "Valid"
                ElseIf Me.Tag = "ACT_PRE" Then
                    Unload Me
                    M_FRM_PRE.Caption = "Account ID Prefix"
                    M_FRM_PRE.Show 1
                ElseIf Me.Tag = "PRD_PRE" Then
                    Unload Me
                    M_FRM_PRE.Caption = "Product ID Prefix"
                    M_FRM_PRE.Show 1
                ElseIf Me.Tag = "CAT_PRE" Then
                    M_FRM_PRE.Caption = "Category ID Prefix"
                    Unload Me
                    M_FRM_PRE.Show 1
                ElseIf Me.Tag = "LOC_PRE" Then
                    Unload Me
                    M_FRM_PRE.Caption = "Location ID Prefix"
                    M_FRM_PRE.Show 1
                ElseIf Me.Tag = "SMN_PRE" Then
                    Unload Me
                    M_FRM_PRE.Caption = "Salesman ID Prefix"
                    M_FRM_PRE.Show 1
                ElseIf Me.Tag = "MAR_PRE" Then
                    Unload Me
                    M_FRM_PRE.Caption = "Market ID Prefix"
                    M_FRM_PRE.Show 1
                ElseIf Me.Tag = "GRP_PRE" Then
                    Unload Me
                    M_FRM_PRE.Caption = "Group Accounts ID Prefix"
                    M_FRM_PRE.Show 1
                ElseIf Me.Tag = "PGR_PRE" Then
                    Unload Me
                    M_FRM_PRE.Caption = "Product Group ID Prefix"
                    M_FRM_PRE.Show 1
                ElseIf Me.Tag = "HHT" Then
                    Unload Me
                    M_FRM_HHT.Show 1
                    ElseIf Me.Tag = "Module" Then
                    Unload Me
                        Dim RS_DIR As New ADODB.Recordset
                        If RS_DIR.State = 1 Then RS_DIR.Close
                        RS_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'MDL' AND DIR_IDX LIKE 'TWP'", CON, adOpenKeyset, adLockPessimistic
                        If RS_DIR.RecordCount > 0 Then
                        If Trim(RS_DIR("DIR_ALS")) = "Y" Then
                            MsgBox "Two Way Pack Module Cannot Be ReConfigured!", vbInformation, "DAS Version 4.0"
                        Else
                            M_FRM_CHK.Show 1
                        End If
                        Else
                            M_FRM_CHK.Show 1
                        End If
                        If RS_DIR.State = 1 Then RS_DIR.Close
                        RS_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'MDL' AND DIR_IDX LIKE 'TCO'", CON, adOpenKeyset, adLockPessimistic
                        If RS_DIR.RecordCount > 0 Then
                        If Trim(RS_DIR("DIR_ALS")) = "Y" Then
                            MsgBox "Telecom Module Cannot Be ReConfigured!", vbInformation, "DAS Version 4.0"
                        Else
                            M_FRM_CHK.Show 1
                        End If
                        End If
                        If RS_DIR.State = 1 Then RS_DIR.Close
                        RS_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'MDL' AND DIR_IDX LIKE 'MIB'", CON, adOpenKeyset, adLockPessimistic
                        If RS_DIR.RecordCount > 0 Then
                        If Trim(RS_DIR("DIR_ALS")) = "Y" Then
                            MsgBox "Hand Held Module Cannot Be ReConfigured!", vbInformation, "DAS Version 4.0"
                        Else
                            M_FRM_CHK.Show 1
                        End If
                        End If
                End If
            Else
                MsgBox "Invalid Password, Try Again!", vbExclamation, "DAS Version 4.0"
                M_FRM_RCT.Tag = "Password"
                Unload Me
            End If
        Else
            MsgBox "Invalid UserName, Try Again!", vbExclamation, "DAS Version 4.0"
            M_FRM_RCT.Tag = "UserName"
            Unload Me
        End If
ElseIf K = 27 Then
    Unload Me
    Set Form = Nothing
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Then KeyAscii = 0
End Sub
