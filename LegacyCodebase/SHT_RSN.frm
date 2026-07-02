VERSION 5.00
Begin VB.Form SHT_RSN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shut Down Restore - Reason"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "SHT_RSN.frx":0000
   LinkTopic       =   "SHT_RSN"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   1020
      TabIndex        =   1
      Top             =   1980
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&No"
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   1980
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   300
      TabIndex        =   0
      Text            =   " "
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   $"SHT_RSN.frx":08CA
      Height          =   675
      Left            =   300
      TabIndex        =   4
      Top             =   1200
      Width           =   4320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   0
      X2              =   4800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter reason for the Improper Shutdown - Mandatory"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   120
      Width           =   3705
   End
End
Attribute VB_Name = "SHT_RSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ShutDown_Write
Unload Me
End Sub

Private Sub Command2_Click()
If Trim(text1) = "" Then
    MsgBox "Please Give Improper ShutDown Reason!", vbInformation, "DAS Version 4.0"
    text1.SetFocus
    Exit Sub
End If
Dim DIR_RS As New ADODB.Recordset
CheckNetWorkLogin = GetValue("DIR_MSG ", "DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'NWL'")
If Trim(CheckNetWorkLogin) = "T" Then
    MsgBox ("Network Login Checked, Proceeding with Login..."), vbInformation, "DAS Version 4.0"
    Command2.Enabled = False
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'IND'"), CON, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveFirst
        DIR_RS("DIR_MSG") = "PROPER SHUTDOWN"
        DIR_RS("DIR_DAT") = Date
        DIR_RS.Update
    End If
    DIR_RS.Close
    Unload Me
Else
    CON.Close
    If Dir(CURDIR & "\DBFIX.BAT") = "" Then
        ShutDown_Write
        MsgBox "DBFIX Batch Utility Not Found!", vbCritical, "DAS Version 4.0"
        CON.Open
        Command2.SetFocus
    Else
        ShutDown_Write
        Shell "DBFIX.BAT", vbMinimizedFocus
        End
    End If
    Command2.Enabled = False
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'IND'"), CON, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveFirst
        DIR_RS("DIR_MSG") = "IMPROPER SHUTDOWN"
        DIR_RS("DIR_DAT") = Date
        DIR_RS.Update
    End If
    DIR_RS.Close
    Unload Me
End If
End Sub

Private Sub Command3_Click()
text1.SetFocus
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K = 13 Or K = 40 Then
    If Trim(text1) = "" Then
        MsgBox "Please Give Improper ShutDown Reason!", vbInformation, "DAS Version 4.0"
        text1.SetFocus
    Else
        Command2.SetFocus
    End If
End If
End Sub
Private Sub ShutDown_Write()
' * Routine to Write Reason for ImProper ShutDown Information into SHUTDOWN.LOG Batch File 'PNR
    Dim DEX As String
    DEX = "C:\DEX"
    If Dir(DEX, vbDirectory) = "" Then
        MkDir DEX
    End If
    Dim FS_SOU, F_VER As Object
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    
    If Dir(App.Path & "\SHUTDOWN.LOG") = "" Then
        Set F_VER = FS_SOU.OpenTextFile(App.Path & "\SHUTDOWN.LOG", 2, True)
        F_VER.WriteLine "REM *** DAS Improper ShutDown Reason ***"
        F_VER.WriteLine "Date " & Date$ & " : Time " & Time$
        F_VER.WriteLine Trim(text1)
        F_VER.Close
    Else
        Set F_VER = FS_SOU.OpenTextFile(App.Path & "\SHUTDOWN.LOG", 8, True)
        F_VER.WriteLine "REM *** DAS Improper ShutDown Reason ***"
        F_VER.WriteLine "Date " & Date$ & " : Time " & Time$
        F_VER.WriteLine Trim(text1)
        F_VER.Close
    End If
    CopyFile App.Path & "\SHUTDOWN.LOG", CURDIR & "\SHUTDOWN.LOG", 0
' * End of Reason Writing Code into ShutDown Batch File Creation 'PNR
End Sub
