VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_EXT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS, Exit to Windows Desktop"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "M_FRM_EXT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   165
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3360
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2085
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   885
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Auto Backup in Progress"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1095
      TabIndex        =   3
      Top             =   960
      Width           =   1755
   End
End
Attribute VB_Name = "M_FRM_EXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CTR, COM_CTR As Integer
Private Sub Command1_Click()
Command2.Enabled = False
Command1.Enabled = False
'Added by PNR for TWP Saving through Control Admin Login Option
        Dim DIR_RS As New ADODB.Recordset
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'MDL' AND DIR_IDX = 'TWP'"), con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            DIR_RS.MoveFirst
            DIR_RS("DIR_MSG") = ""
            DIR_RS.Update
        End If
' Added by PNR for TWP Saving through Control Admin Login Option
        If DIR_RS.State = 1 Then DIR_RS.Close
        DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'IND'"), con, adOpenKeyset, adLockPessimistic
        If DIR_RS.RecordCount > 0 Then
            DIR_RS.MoveFirst
            If Trim(DIR_RS("DIR_MSG")) = "IMPROPER SHUTDOWN" Then
                MsgBox ("Please reset the Improper Shutdown flag in System-ShutdownStatus!"), vbInformation, "DAS Version 4.0"
            Else
                DIR_RS("DIR_MSG") = "PROPER SHUTDOWN"
            End If
            DIR_RS("DIR_DAT") = Date
            DIR_RS.Update
        Else
            DIR_RS.AddNew
            DIR_RS("DIR_IDY") = "SYS"
            DIR_RS("DIR_IDX") = "IND"
            DIR_RS("DIR_ALS") = "T"
            DIR_RS("DIR_DAT") = Date
            DIR_RS("DIR_MSG") = "PROPER SHUTDOWN"
            DIR_RS.Update
        End If
LOGOUT
' * AutoBackup While Exit
Dim SPATH As String
Dim DPATH As String
CTR = 8
If querycheck("DIR", "DIR_MSG", "T", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'ABK'") = True Then
    If DIR_RS.State = 1 Then DIR_RS.Close
    DIR_RS.Open ("SELECT * FROM DIR WHERE DIR_IDY = 'SYS' AND DIR_IDX = 'IND'"), con, adOpenKeyset, adLockPessimistic
    If DIR_RS.RecordCount > 0 Then
        DIR_RS.MoveFirst
        If Trim(DIR_RS("DIR_MSG")) = "PROPER SHUTDOWN" Then
            Me.Height = 2655
            Me.AutoRedraw = True
            Me.Refresh
            Label1.Visible = True
            ProgressBar1.Visible = True
            Me.MousePointer = vbHourglass
            AUTOBKP = CURDIR & "\AUTOBKP"
            If Dir(AUTOBKP, vbDirectory) = "" Then
                MkDir AUTOBKP
            End If
            INC_PCB
            Dim FS As Object, f As Object, F1 As Object, FN As Object
            Set FS = CreateObject("Scripting.FileSystemObject")
            SPATH = CURDIR & "\???.DBF"
            DPATH = CURDIR & "\AUTOBKP"
            FS.CopyFile SPATH, DPATH
            INC_PCB
            SPATH = CURDIR & "\EFF.*"
            DPATH = CURDIR & "\AUTOBKP"
            FS.CopyFile SPATH, DPATH
            INC_PCB
            SPATH = CURDIR & "\*.CDX"
            DPATH = CURDIR & "\AUTOBKP"
            FS.CopyFile SPATH, DPATH
            INC_PCB
            SPATH = CURDIR & "\*.FPT"
            DPATH = CURDIR & "\AUTOBKP"
            FS.CopyFile SPATH, DPATH
            INC_PCB
            SPATH = App.Path & "\*.RP1"
            DPATH = CURDIR & "\AUTOBKP"
            FS.CopyFile SPATH, DPATH
            INC_PCB
            SPATH = App.Path & "\*.RRW"
            DPATH = CURDIR & "\AUTOBKP"
            FS.CopyFile SPATH, DPATH
            INC_PCB
            CopyFile App.Path & "\IND.EXE", DPATH & "\IND.EXE", 0
            CopyFile App.Path & "\CND.EXE", DPATH & "\CND.EXE", 0
            CopyFile App.Path & "\RRUNTIME.EXE", DPATH & "\RRUNTIME.EXE", 0
            CopyFile App.Path & "\RRUN.EXE", DPATH & "\RRUN.EXE", 0
            CopyFile App.Path & "\RR.CNF", DPATH & "\RR.CNF", 0
            CopyFile App.Path & "\PRI.EXE", DPATH & "\PRI.EXE", 0
            CopyFile App.Path & "\PRB.EXE", DPATH & "\PRB.EXE", 0
            INC_PCB
            'MsgBox ("AutoBackup Completed Successfully !"), vbInformation, "DAS Version 4.0"
            Me.MousePointer = vbDefault
        Else
            MsgBox ("AutoBackup UnSuccessful ! Improper Shutdown Encountered"), vbInformation, "DAS Version 4.0"
        End If
    End If
End If
' * End of AutoBackup 'PNR
' * Backup Date Stamp and Path Record - added by PNR
    Dim RST_DIR As New ADODB.Recordset
    If RST_DIR.State = 1 Then RST_DIR.Close
    RST_DIR.Open "SELECT * FROM DIR WHERE DIR_IDY LIKE 'BKP' .AND. DIR_IDX LIKE 'LOG'", con, adOpenKeyset, adLockPessimistic
    If RST_DIR.EOF = False Then
        RST_DIR.MoveFirst
        If Trim(DPATH) <> "" Then
            RST_DIR("DIR_MSG") = Trim(DPATH)
            RST_DIR("DIR_DAT") = Format(Date, "MM/DD/YYYY")
        End If
        RST_DIR.Update
    Else
        RST_DIR.AddNew
        RST_DIR("DIR_IDY") = "BKP"
        RST_DIR("DIR_IDX") = "LOG"
        RST_DIR("DIR_ALS") = ""
        RST_DIR("DIR_MSG") = Trim(DPATH)
        RST_DIR("DIR_DAT") = Format(Date, "MM/DD/YYYY")
        RST_DIR.Update
    End If
' * Backup Date and Path Record - added by PNR
Unload Me
Set Form = Nothing

' * Routine to Create Proper ShutDown Information Batch File 'PNR

Dim DEX As String
DEX = "C:\DEX"
If Dir(DEX, vbDirectory) = "" Then
    MkDir DEX
End If
Dim FS_SOU, F_VER As Object
Set FS_SOU = CreateObject("Scripting.filesystemobject")

If Dir(App.Path & "\SHUTDOWN.LOG") = "" Then
    Set F_VER = FS_SOU.OpenTextFile(App.Path & "\SHUTDOWN.LOG", 2, True)
    F_VER.WriteLine "REM *** DAS Normal ShutDown ***"
    F_VER.WriteLine "Date " & Date$ & " : Time " & Time$
    F_VER.Close
Else
    Set F_VER = FS_SOU.OpenTextFile(App.Path & "\SHUTDOWN.LOG", 8, True)
    F_VER.WriteLine "REM *** DAS Normal ShutDown ***"
    F_VER.WriteLine "Date " & Date$ & " : Time " & Time$
    F_VER.Close
End If
CopyFile App.Path & "\SHUTDOWN.LOG", CURDIR & "\SHUTDOWN.LOG", 0
' * End of ShutDown Batch File Creation 'PNR

' * Begin of Indexing While Exit 'PNR
'''If querycheck("DIR", "DIR_ALS", "I", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'") = True Then
'''        CON.Close
'''        PNL.Close
'''        Set CON = Nothing
'''        Set PNL = Nothing
'''        Timer1.Enabled = True
'''        'Code Added by PNR to make users index at the time of Application Loading!
'''        'If MsgBox("Index DAS Database ? - Recommended Maintenance Task.. Takes Time!", vbYesNo + vbQuestion, "DAS Version 4.0") = vbYes Then
'''            ChDir CURDIR
'''            ' Dir(CURDIR & "\IND.EXE") = "" Or Dir(CURDIR & "\CND.EXE") = "" Or
'''            If Dir(CURDIR & "\EXIT.BAT") = "" Then
'''                MsgBox "Cannot Process Request, EXIT.BAT Not Found in Folder " & Trim(CURDIR), vbInformation, "DAS Version 4.0"
'''                End
'''            Else
''''                CON.Close
''''                PNL.Close
''''                Set CON = Nothing
''''                Set PNL = Nothing
'''                'Shell CURDIR & "\IND.EXE", vbMaximizedFocus
'''                On Error Resume Next
'''                Shell CURDIR & "\EXIT.BAT", vbMaximizedFocus
'''                End
'''            End If
'''        'End If
'''Else
'''        CON.Close
'''        PNL.Close
'''        Set CON = Nothing
'''        Set PNL = Nothing
'''End If

' * End of Indexing While Exit 'PNR

'''CON.Close
'''PNL.Close
'''Set CON = Nothing
'''Set PNL = Nothing
'''Timer1.Enabled = True

ChDir CURDIR
' Dir(CURDIR & "\IND.EXE") = "" Or Dir(CURDIR & "\CND.EXE") = "" Or
'If querycheck("DIR", "DIR_ALS", "I", " AND DIR_IDY LIKE 'SYS' AND DIR_IDX LIKE 'IND'") = True Then
    If Dir(CURDIR & "\EXIT.BAT") = "" Then
        MsgBox "Cannot Process Request, EXIT.BAT Not Found in Folder " & Trim(CURDIR), vbInformation, "DAS Version 4.0"
        End
    Else
        con.Close
        PNL.Close
        Set con = Nothing
        Set PNL = Nothing
        Timer1.Enabled = True
    
        'Shell CURDIR & "\IND.EXE", vbMaximizedFocus
        On Error Resume Next
        Shell CURDIR & "\EXIT.BAT", vbHide
        End
    End If
'End If     'System, Specify Defaults value ignored as per PC requirement for compulsory running of EXIT.BAT
End
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
Me.Height = 1440
Me.Width = 4035
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Shell CURDIR & "\IND.EXE", vbHide
End Sub
Public Sub INC_PCB()
COM_CTR = COM_CTR + 1
ProgressBar1.Value = (100 / CTR) * COM_CTR
End Sub
