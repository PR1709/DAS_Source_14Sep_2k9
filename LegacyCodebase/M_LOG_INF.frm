VERSION 5.00
Begin VB.Form M_LOG_INF 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "M_LOG_INF.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1748
      TabIndex        =   19
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   7
      Text            =   " "
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   6
      Text            =   " "
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   4
      Text            =   " "
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   3
      Text            =   " "
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   165
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   15
      Top             =   1170
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   13
      Top             =   2610
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   12
      Top             =   2130
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   11
      Top             =   1650
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Last Invoice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   645
      TabIndex        =   5
      Top             =   2610
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Last Day Closure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   315
      TabIndex        =   2
      Top             =   1650
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Last Van Challan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   315
      TabIndex        =   1
      Top             =   2130
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Last Backup Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   1170
      Width           =   1290
   End
End
Attribute VB_Name = "M_LOG_INF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Check_BackupDCL_Status()
' Backup Checkup and Message Module - added by PNR
Dim RST_DIR As New ADODB.Recordset
If RST_DIR.State = 1 Then RST_DIR.Close
RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'BKP' .AND. DIR_IDX LIKE 'LOG'", con, adOpenKeyset, adLockPessimistic
If RST_DIR.EOF = True Then
    RST_DIR.AddNew
    RST_DIR("DIR_IDY") = "BKP"
    RST_DIR("DIR_IDX") = "LOG"
    RST_DIR("DIR_ALS") = ""
    RST_DIR("DIR_MSG") = ""
    RST_DIR("DIR_DAT") = Format(Date, "MM/DD/YYYY")
    RST_DIR.Update
End If
BKP_DAYS = DateDiff("D", RST_DIR("DIR_DAT"), Date)
'&H000000FF&
'&H00FFFF80&
If Trim(RST_DIR("DIR_MSG")) <> "" Then
    If Format(Date, "DD/MM/YYYY") <> Format(RST_DIR("DIR_DAT"), "DD/MM/YYYY") Then
        Text1 = Format(RST_DIR("DIR_DAT"), "DD/MM/YYYY")
        Text5 = DateDiff("D", RST_DIR("DIR_DAT"), Date)
        BKP_DAYS = DateDiff("D", RST_DIR("DIR_DAT"), Date)
        If BKP_DAYS > 2 Then
            Text1.BackColor = &HFF&
            Text5.BackColor = &HFF&
        Else
            Text1.BackColor = &HFFFF80
            Text5.BackColor = &HFFFF80
        End If
'        MsgBox "Last Data Backup Date : " & Format(RST_DIR("DIR_DAT"), "DD/MM/YYYY") & " in Folder " & Trim(RST_DIR("DIR_MSG")) & " for " & STR(Val(BKP_DAYS)) & " Days!", vbInformation, "DAS Version 4.0"
    End If
    '"You Didn't Backup for " & Abs(Val(BKP_DAYS)) & " Days!"
Else
        Text1 = "NA"
        Text5 = "NA"
        Text1.BackColor = &HFF&
        Text5.BackColor = &HFF&
        'Text1 = Format(Date, "DD/MM/YYYY")
        'Text5 = DateDiff("D", Date, Date)
'    MsgBox "No Data Backup Found! Take Backup From System -> Database -> Complete Backup Menu Option!", vbInformation, "DAS Version 4.0"
End If
' Backup Checkup and Message Module - added by PNR
If RST_DIR.State = 1 Then RST_DIR.Close
RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'DAY' .AND. alltrim(DIR_MSG) LIKE 'Y' ORDER BY DIR_DAT", con, adOpenKeyset, adLockPessimistic
If RST_DIR.EOF = False Then
    RST_DIR.MoveLast
    Text3 = Format(RST_DIR("DIR_DAT"), "DD/MM/YYYY")
    Text6 = DateDiff("D", RST_DIR("DIR_DAT"), Date)
    If Val(Text6) > 7 Then
        Text3.BackColor = &HFF&
        Text6.BackColor = &HFF&
    Else
        Text3.BackColor = &HFFFF80
        Text6.BackColor = &HFFFF80
    End If
Else
    Text3 = "NA"
    Text6 = "NA"
    End If
End Sub
Private Sub JUST_MADE_Invoice_Click()
Dim INL_JMD As New ADODB.Recordset
If INL_JMD.State = 1 Then INL_JMD.Close
INL_JMD.Open "SELECT * FROM INV WHERE TXN_TYP = 'SAL' ORDER BY INV_DAT", con, adOpenKeyset, adLockPessimistic
If INL_JMD.RecordCount > 0 Then
    INL_JMD.MoveLast
    Text4 = Format(INL_JMD("INV_DAT"), "DD/MM/YYYY")
    Text8 = DateDiff("D", INL_JMD("INV_DAT"), Date)
Else
    Text4 = Format(Date, "DD/MM/YYYY")
    Text8 = DateDiff("D", Date, Date)
End If
INL_JMD.Close
End Sub
Private Sub JUST_MADE_VanChallan_Click()
Dim INL_JMD As New ADODB.Recordset
If INL_JMD.State = 1 Then INL_JMD.Close
INL_JMD.Open "SELECT * FROM VAN ORDER BY INV_DAT", con, adOpenKeyset, adLockPessimistic
If INL_JMD.RecordCount > 0 Then
    INL_JMD.MoveLast
    Text2 = Format(INL_JMD("INV_DAT"), "DD/MM/YYYY")
    Text7 = DateDiff("D", INL_JMD("INV_DAT"), Date)
Else
    Text2 = Format(Date, "DD/MM/YYYY")
    Text7 = DateDiff("D", Date, Date)
End If
INL_JMD.Close
End Sub
Private Sub Form_Activate()
Dim DIR_RS As New ADODB.Recordset
If DIR_RS.State = 1 Then DIR_RS.Close
DIR_RS.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TT1'", con, adOpenKeyset, adLockPessimistic
If DIR_RS.RecordCount > 0 Then
    DIR_RS.MoveFirst
    Label10.Caption = Trim(DIR_RS!DIR_MSG)
    Label10.Refresh
End If
If DIR_RS.State = 1 Then DIR_RS.Close
DIR_RS.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'TT2'", con, adOpenKeyset, adLockPessimistic
If DIR_RS.RecordCount > 0 Then
    DIR_RS.MoveFirst
    Label9.Caption = Trim(DIR_RS!DIR_MSG)
    Label9.Refresh
End If
If DIR_RS.State = 1 Then DIR_RS.Close
DIR_RS.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'WDS' AND DIR_IDX LIKE 'IDY'", con, adOpenKeyset, adLockPessimistic
If DIR_RS.RecordCount > 0 Then
    DIR_RS.MoveFirst
    Label11.Caption = Trim(DIR_RS!DIR_ALS)
    Label11.Refresh
End If
DIR_RS.Close
M_LOG_INF.Caption = "Actvty Log " & Trim(SEL_CMP)
Command1.SetFocus
End Sub

Private Sub Form_Load()
Check_BackupDCL_Status
JUST_MADE_Invoice_Click
JUST_MADE_VanChallan_Click
End Sub


