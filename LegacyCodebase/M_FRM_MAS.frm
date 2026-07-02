VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form M_FRM_MAS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Import"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   Icon            =   "M_FRM_MAS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Select"
      Height          =   615
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Width           =   3015
      Begin VB.OptionButton opt_efx 
         Caption         =   "Efflinx"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton opt_eff 
         Caption         =   "DAS"
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5160
      Width           =   855
   End
   Begin VB.CheckBox All 
      Caption         =   "Complete Data&base"
      Height          =   195
      Left            =   4800
      TabIndex        =   8
      Top             =   960
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Select Source Folder"
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4095
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   3855
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblvalidfolder 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   3150
         Width           =   3855
      End
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   4800
      Top             =   -600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select File Information for Data Import"
      Height          =   3495
      Left            =   4560
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
      Begin VB.CheckBox SCH_CHK 
         Caption         =   "Sc&hemes"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   3180
         Width           =   1455
      End
      Begin VB.CheckBox CAT_CHK 
         Caption         =   "Ca&tegories"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox CSR_CHK 
         Caption         =   "Preset &Rates"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   2410
         Width           =   1215
      End
      Begin VB.CheckBox PRD_CHK 
         Caption         =   "&Products"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox MAR_CHK 
         Caption         =   "&Markets"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1140
         Width           =   975
      End
      Begin VB.CheckBox SMN_CHK 
         Caption         =   "&Salesman"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   885
         Width           =   1095
      End
      Begin VB.CheckBox LOC_CHK 
         Caption         =   "&Locations"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   630
         Width           =   1095
      End
      Begin VB.CheckBox PGR_CHK 
         Caption         =   "Pro&duct Categories"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1900
         Width           =   1695
      End
      Begin VB.CheckBox GRP_CHK 
         Caption         =   "&Group Accounts"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1650
         Width           =   1455
      End
      Begin VB.CheckBox ACT_CHK 
         Caption         =   "&Accounts"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1390
         Width           =   975
      End
      Begin VB.CheckBox ASM_CHK 
         Caption         =   "Ass&ets"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2670
         Width           =   1215
      End
      Begin VB.CheckBox SHL_CHK 
         Caption         =   "Shells &n Packs"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2920
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   6360
      TabIndex        =   2
      Top             =   4920
      Width           =   2295
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   2520
      Width           =   495
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   1200
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importing Data .. Please Wait"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2760
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Import in progress.. Please wait.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   5805
      Visible         =   0   'False
      Width           =   3945
   End
End
Attribute VB_Name = "M_FRM_MAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CN As New ADODB.Connection
'Dim CN1 As New ADODB.Connection
'Dim cn2 As New ADODB.Connection
'Dim rs2 As New ADODB.Recordset
'Dim rs3 As New ADODB.Recordset
'Dim db As Database
'Dim RS As New ADODB.Recordset
'Dim tb As TableDef
'Dim DB1 As Database
'Dim RS1 As New ADODB.Recordset
'Dim tb1 As TableDef
'Dim dbs As Database
Dim CNX As New ADODB.Connection
Dim CNG As New ADODB.Connection
Dim RSX As New ADODB.Recordset
Dim RSG As New ADODB.Recordset
Dim RST As New ADODB.Recordset
Dim RST1 As New ADODB.Recordset
Dim rid As New ADODB.Recordset
Dim tns As New ADODB.Recordset
Dim pr As String, pr1 As String, found As Integer
Dim f As Field
Dim I As Integer, n As Integer, id As Integer, id1 As Integer, id2 As Integer, FF As Integer
Dim vr As String

Private Sub ACT_CHK_Click()
'If opt_efx.Value = True Then
'If CHECKTABLE("ACT") = 0 Then
'MsgBox "Table does not exist!", vbOKOnly + vbInformation, "Data Import"
'ACT_CHK.Value = 0
'Else
'ACT_CHK.Value = 1
'End If
'End If
If ACT_CHK.Value = 1 Then
list1.AddItem "act"
Else
For I = 0 To list1.ListCount - 1
If list1.List(I) = "act" Then
list1.RemoveItem I
Exit For
End If
Next
End If
End Sub

Private Sub All_Click()
If All.Value = 1 Then
With list1
    .AddItem "aca"
    .AddItem "asm"
    .AddItem "ast"
    .AddItem "att"
    .AddItem "sic"
    .AddItem "ord"
    .AddItem "cur"
    .AddItem "wek"
    .AddItem "usr"
    .AddItem "qps"
    .AddItem "opa"
    .AddItem "ops"
    .AddItem "rdy"
    .AddItem "yer"
    .AddItem "typ"
    .AddItem "snv"
    .AddItem "shl"
    .AddItem "scl"
    .AddItem "tgc"
    .AddItem "tgb"
    .AddItem "tgs"
    .AddItem "tgt"
    .AddItem "sch"
    .AddItem "brd"
    .AddItem "irt"
    .AddItem "lds"
    
    ' Complete Transaction Data - Complete Data
    .AddItem "act"
    .AddItem "cat"
    .AddItem "csr"
    .AddItem "dir"
    .AddItem "grp"
    .AddItem "idy"
    .AddItem "inl"
    .AddItem "inv"
    .AddItem "lev"
    .AddItem "lnv"
    .AddItem "loc"
    .AddItem "mar"
    .AddItem "pgr"
    .AddItem "prd"
    .AddItem "pre"
    .AddItem "prt"
    .AddItem "rtp"
    .AddItem "smn"
    .AddItem "van"
    .AddItem "vch"
    
    ' Only Master Data
    
'    .AddItem "act"
'    .AddItem "cat"
'    .AddItem "csr"
'    .AddItem "dir"
'    .AddItem "grp"
'    .AddItem "idy" 'When master data import is selected all the fields where transaction record number should be made to zero
'    .AddItem "loc"
'    .AddItem "mar"
'    .AddItem "pgr"
'    .AddItem "prd"
'    .AddItem "pre"
'    .AddItem "prt"
'    .AddItem "smn"
End With
Else
    list1.CLEAR
End If
End Sub

Private Sub ASM_CHK_Click()
If ASM_CHK.Value = 1 Then
list1.AddItem "asm"
list1.AddItem "att"
list1.AddItem "aca"
list1.AddItem "ast"
Else
For I = 0 To list1.ListCount - 1
If list1.List(I) = "asm" Then
list1.RemoveItem I
Exit For
End If
Next
For I = 0 To list1.ListCount - 1
If list1.List(I) = "att" Then
list1.RemoveItem I
Exit For
End If
Next
For I = 0 To list1.ListCount - 1
If list1.List(I) = "aca" Then
list1.RemoveItem I
Exit For
End If
Next
For I = 0 To list1.ListCount - 1
If list1.List(I) = "ast" Then
list1.RemoveItem I
Exit For
End If
Next
End If
End Sub

Private Sub CAT_CHK_Click()
If CAT_CHK.Value = 1 Then
list1.AddItem "cat"
Else
For I = 0 To list1.ListCount - 1
If list1.List(I) = "cat" Then
list1.RemoveItem I
Exit For
End If
Next
End If
End Sub

Private Sub Command1_Click()
On Error GoTo errh
If opt_eff.Value = True Then
'pb.Max = List1.ListCount + 10
tim = Time$
If Dir1.Path = App.Path Then
    MsgBox "Current Folder Path Selected! Invalid Folder Path to import File Data", vbCritical, "DAS Version 4.0"
    Exit Sub
End If
If Len(Dir1.Path) <> 3 Then
    If Dir(Dir1.Path & "\EFF.DBC") = "" Then
        MsgBox "Not A Valid DAS Version 4.0 Database Folder Path", vbCritical, "DAS Version 4.0"
        Exit Sub
    End If
Else
    If Dir(Dir1.Path & "EFF.DBC") = "" Then
        MsgBox "Not A Valid DAS Version 4.0 Database Folder Path", vbCritical, "DAS Version 4.0"
        Exit Sub
    End If
End If
If MsgBox("This Action Will Delete Existing Data ! Continue?", vbYesNo + vbQuestion, "DAS Version 4.0") = vbNo Then
    Exit Sub
Else
    If MsgBox("Confirm Data Import From " & Dir1.Path, vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then Exit Sub
End If
'pb.Visible = True
Label4.Visible = True
Me.MousePointer = vbHourglass
Label3.Visible = True
Label3.Refresh
'strattributes = "DBQ=" & Dir1.Path & "\Eff.dbc" & vbCr & "Description=Effmcg Back Up" & vbCr & "SourceType=DBC" & vbCr & "Deleted=Yes"
'DBEngine.RegisterDatabase "EFFc", "Microsoft Visual Foxpro Driver", True, strattributes
'CN.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=EFFc"
If cn.State = 1 Then cn.Close
cn.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & Dir1.Path & "\Eff.dbc;SourceType=DBc;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
'CN.Open "EFFc"
'strattributes = "DBQ=" & CURDIR & "\Eff.dbc" & vbCr & "Description=Effmcg Back Up" & vbCr & "SourceType=DBC" & vbCr & "Deleted=Yes"
'DBEngine.RegisterDatabase "EFF", "Microsoft Visual Foxpro Driver", True, strattributes
'CN1.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=EFF"
'MsgBox CURDIR
'strattributes = "DBQ=" & CURDIR & vbCr & "Description=Effmcg Back Up" & vbCr & "SourceType=DBf" & vbCr & "Deleted=Yes"
'DBEngine.RegisterDatabase "pnl1", "Microsoft Visual Foxpro Driver", True, strattributes
If CN1.State = 1 Then CN1.Close
CN1.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & "\Eff.dbc;SourceType=DBc;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
If DB.State = 1 Then DB.Close
DB.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & Dir1.Path & "\Eff.dbc;SourceType=DBc;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
'Set db = OpenDatabase(Dir1.Path + "\Eff.dbc", False, False)
'Set db = OpenDatabase("", dbDriverNoPrompt, False, "ODBC;DSN=EFFC")
'Set DB1 = OpenDatabase("", dbDriverNoPrompt, False, "ODBC;DSN=EFF")
'Set DB1 = OpenDatabase(CURDIR + "\Eff.dbc", False, False)
If DB1.State = 1 Then DB1.Close
DB1.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & "\Eff.dbc;SourceType=DBc;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
If cn2.State = 1 Then cn2.Close
cn2.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & ";SourceType=DBf;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
'Set dbs = OpenDatabase("", dbDriverNoPrompt, False, "ODBC;DSN=PNL1")
'Set dbs = OpenDatabase(CURDIR, False, False)
If dbs.State = 1 Then dbs.Close
dbs.Open "Driver={Microsoft Visual FoxPro Driver};UID=;PWD=;SourceDB=" & CURDIR & ";SourceType=DBf;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"
'imp_sta.Min = 0
'imp_sta.Max = db.TableDefs.Count - 1
Set tns = DB.OpenSchema(adSchemaTables)
id = 0
For I = 0 To tns.RecordCount - 1 'db.TableDefs.Count - 1
'Set tb = TNS!TABLE_NAME 'db.TableDefs(i)
For K = 0 To LIST2.ListCount - 1
If tns!TABLE_NAME = LIST2.List(K) Then
id = id + 1
End If
Next
tns.MoveNext
Next
'If id = 0 Then
'MsgBox "Version 2.02"
'ElseIf id = 1 Then
'MsgBox "Version 3.01"
'ElseIf id = 2 Then
'MsgBox "Version 4.00"
'End If
'Me.MousePointer = vbHourglass
'strattributes = "DBQ=" & Dir1.Path & "\Eff.dbc" & vbCr & "Description=Effmcg Back Up" & vbCr & "SourceType=DBC" & vbCr & "Deleted=Yes"
'DBEngine.RegisterDatabase "EFFBKP", "Microsoft Visual Foxpro Driver", True, strattributes
'CN.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=EFF1"
''CN.Open "DSN=EFFBKP"
'If CN.State = 1 Then CN.Close
'CN.Open "EFFBKP"
'strattributes = "DBQ=" & App.Path & "\Eff.dbc" & vbCr & "Description=Effmcg Back Up" & vbCr & "SourceType=DBC" & vbCr & "Deleted=Yes"
'DBEngine.RegisterDatabase "EFF", "Microsoft Visual Foxpro Driver", True, strattributes
'CN1.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=EFF"
'If CN1.State = 1 Then CN1.Close
'CN1.Open "EFF"
'Set db = OpenDatabase("", dbDriverNoPrompt, False, "ODBC;DSN=EFFBKP")
'Set DB1 = OpenDatabase("", dbDriverNoPrompt, False, "ODBC;DSN=EFF")
''Set db = OpenDatabase(Dir1.Path + "\eff.dbc", dbDriverNoPrompt, False, "FoxPro DBC;")
''Set DB1 = OpenDatabase("c:\importdata\eff.dbc", dbDriverNoPrompt, False, "FoxPro DBC;")
'If Dir1.Path = App.Path Then
'    MsgBox "Current Folder Path Selected! Invalid Folder Path to import File Data", vbCritical, "DAS Version 4.0"
'    Exit Sub
'End If
'If Dir(Dir1.Path & "\EFF.DBC") = "" Then
'    MsgBox "Not A Valid DAS Version 4.0 Database Folder Path", vbCritical, "DAS Version 4.0"
'    Exit Sub
'End If
'For t = 0 To TNS.RecordCount - 1 'db.TableDefs.Count - 1
tns.MoveFirst
While Not tns.EOF
'Set tb = db.TableDefs(t)
    For n = 0 To list1.ListCount - 1
        If Trim(tns!TABLE_NAME) = list1.List(n) Then
            If rs.State = 1 Then rs.Close
            If RS1.State = 1 Then RS1.Close
            CN1.Execute "delete from " & tns!TABLE_NAME
            rs.Open "select * from " & tns!TABLE_NAME, cn, adOpenDynamic, adLockPessimistic
            RS1.Open "select * from " & tns!TABLE_NAME, CN1, adOpenDynamic, adLockPessimistic
            If Trim(tns!TABLE_NAME) = "usr" Then
                While Not rs.EOF
                    If Len(Trim(rs("USR_PWD"))) > 8 Then
                        tpwd = Trim(rs("usr_pwd"))
                    Else
                        tpwd = EPWD.encPwd(Trim(rs("usr_pwd")))
                    End If
                    RS1.AddNew
                    RS1(0) = rs(0)
                    RS1(1) = Trim(tpwd)
                    RS1(2) = rs(2)
                    RS1(3) = rs(3)
                    RS1.Update
                    rs.MoveNext
                Wend
            ElseIf Trim(tns!TABLE_NAME) = "act" Then
                While Not rs.EOF
                    RS1.AddNew
                        For I = 0 To rs.Fields.COUNT - 1
                            FLD = rs.Fields(I).name
                            If Trim(FLD) = "cst_typ" Then
                                If Trim(rs(FLD)) = "" Then
                                    RS1(FLD) = "Retailer"
                                Else
                                    RS1(FLD) = rs(FLD)
                                End If
                            ElseIf Trim(FLD) = "dbt_rap" Then
                                If Trim(rs(FLD)) = "" Then
                                    RS1(FLD) = "B"
                                Else
                                    RS1(FLD) = rs(FLD)
                                End If
                            Else
                                RS1(FLD) = rs(FLD)
                            End If
                        Next
                    RS1.Update
                    rs.MoveNext
                    RS1.MoveNext
                Wend
            Else
                While Not rs.EOF
                    RS1.AddNew
                        For I = 0 To rs.Fields.COUNT - 1
                            FLD = rs.Fields(I).name
                            RS1(FLD) = rs(FLD)
                        Next
                    RS1.Update
                    rs.MoveNext
                    RS1.MoveNext
                Wend
            End If
                Exit For
            'pb.Value = pb.Value + 2
        End If
    Next
'    imp_sta.Value = t
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    tns.MoveNext
Wend
If Dir(Dir1.Path + "\rdy.dbf") = "" Then
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    rs.Open "select act_idy,smn_idy,rou_idy from act where act_typ='SAL'", CON, adOpenKeyset, adLockPessimistic
    RS1.Open "select * from rdy where act_idy=''", CON, adOpenKeyset, adLockPessimistic
    While Not rs.EOF
        For I = 1 To 7
            RS1("act_idy") = rs("act_idy")
            RS1("day_idy") = I
            RS1("smn_idy") = rs("smn_idy")
            RS1("rou_idy") = rs("rou_idy")
            RS1("seq_num") = 0
            RS1("rou_dst") = 0
            RS1("tim_inn") = 0
            RS1("tim_out") = 0
        Next
        RS1.MoveNext
    Wend
End If
If id = 2 Then
If All.Value = 1 Then
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    CN1.Execute "delete from pnl"
rs.Open "select * from pnl", cn, adOpenDynamic, adLockPessimistic
RS1.Open "select * from pnl", CN1, adOpenDynamic, adLockPessimistic
'Set tb = dbs.TableDefs(3)
While Not rs.EOF
RS1.AddNew
For I = 0 To rs.Fields.COUNT - 1
    FLD = rs.Fields(I).name
    RS1(FLD) = rs(FLD)
Next
    RS1.Update
    rs.MoveNext
    'RS1.MoveNext
Wend
'imp_sta.Value = imp_sta.Value + 1
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    CN1.Execute "delete from pnv"
rs.Open "select * from pnv", cn, adOpenDynamic, adLockPessimistic
RS1.Open "select * from pnv", CN1, adOpenDynamic, adLockPessimistic
'Set tb = dbs.TableDefs(4)
While Not rs.EOF
RS1.AddNew
For I = 0 To rs.Fields.COUNT - 1
    FLD = rs.Fields(I).name
    RS1(FLD) = rs(FLD)
Next
    RS1.Update
    rs.MoveNext
    RS1.MoveNext
Wend
'imp_sta.Value = imp_sta.Value + 1
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
End If
'pb.Value = pb.Value + 2
End If
'MsgBox "TWO"
If id = 0 Or id = 1 Then
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    rs.Open "SELECT * FROM PRD", CN1, adOpenDynamic, adLockPessimistic
    RS1.Open "SELECT * FROM BRD", CN1, adOpenDynamic, adLockPessimistic
    'MsgBox "OK"
    'Set tb = DB1.TableDefs(3)
    CN1.Execute "delete from brd"
    While Not rs.EOF
        RS1.AddNew
        For I = 0 To rs.Fields.COUNT
            FLD = rs.Fields(I).name
            RS1(FLD) = rs(FLD)
        Next
        RS1!Bat_nme = ""
        RS1("PRD_PDR") = rs("pur_rat")
        RS1("prd_sdr") = rs("sal_rat")
        RS1("prd_mrp") = rs("mrp_rat")
        RS1("bat_idy") = "BAT" & Right(rs("prd_idy"), 7)
        RS1.Update
        rs.MoveNext
    Wend
    CN1.Execute "UPDATE IDY SET IDY_BAT=''"
    CN1.Execute "UPDATE IDY SET IDY_BAT=IDY_PRD"
    'MsgBox "OK"
  'imp_sta.Value = imp_sta.Value + 1
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    RS1.Open "select * from brd", CN1, adOpenDynamic, adLockPessimistic
    While Not RS1.EOF
        CN1.Execute "update inv set bat_idy='" & Trim(RS1("bat_idy")) & "'" & " where prd_idy like '" & Trim(RS1("prd_idy")) & "'"
        RS1.MoveNext
    Wend
    'MsgBox "OK UPTO"
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    'MsgBox "OK"
    'pb.Value = pb.Value + 1
End If
'MsgBox "two step"
If id = 0 Then
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    If rs2.State = 1 Then rs2.Close
    rs.Open "SELECT * FROM INL", CN1, adOpenDynamic, adLockPessimistic
    'MsgBox "OK"
    'rs2.Open "select * from pnl", CN1, adOpenDynamic, adlockpessimistic
    'CN1.Execute "delete from pnl"
    'If RS1.State = 1 Then RS1.Close
    CN1.Execute "delete from irt"
    RS1.Open "SELECT * FROM IRT", CN1, adOpenDynamic, adLockPessimistic
    'MsgBox "IRT"
    'Set tb = DB1.TableDefs(15)
    'RS.MoveFirst
    While Not rs.EOF
        RS1.AddNew
        For I = 0 To rs.Fields.COUNT - 1
        FLD = rs.Fields(I).name
        RS1(FLD) = rs(FLD)
        Next
'        RS1("inv_idy") = rs("inv_idy")
'        RS1("inv_dat") = rs("inv_dat")
'        RS1("due_dat") = rs("due_dat")
'        RS1("net_amt") = rs("net_amt")
'        RS1("txn_typ") = rs("txn_typ")
'        RS1("act_idy") = rs("act_idy")
'        RS1("smn_idy") = rs("smn_idy")
        RS1("PTY_IDY") = rs("CST_IDY")
        RS1.Update
'    Set tb = dbs.TableDefs(3)
'    If RS1.State = 1 Then RS1.Close
'    RS1.Open "SELECT * FROM pnl", CN1, adOpenDynamic, adlockpessimistic
'        RS1.AddNew
'        For i = 0 To rs.Fields.Count - 1
'        fld = rs.Fields(i).name
'        RS1(fld) = rs(fld)
'        Next
'        RS1.Update
        rs.MoveNext
    Wend
    If rs.State = 1 Then rs.Close
    If RS1.State = 1 Then RS1.Close
    'MsgBox "IRT OVER"
    'pb.Value = pb.Value + 1
End If
If id = 0 Or id = 1 Then
inltopnl
'MsgBox "INLTOPNL"
'If rs.State = 1 Then rs.Close
'If RS1.State = 1 Then RS1.Close
'rs.Open "select * from inl", CN1, adOpenKeyset, adlockpessimistic
'MsgBox "Number of records in inl " & rs.RecordCount
'CN1.Execute "delete from pnl"
'Set tb = dbs.TableDefs(3)
'RS1.Open "select * from pnl", CN1, adOpenDynamic, adlockpessimistic
'rs.MoveFirst
'While Not rs.EOF
'RS1.AddNew
'For i = 0 To rs.Fields.Count - 1
'fld = rs.Fields(i).name
'RS1(fld) = rs(fld)
'Next
'RS1.Update
'RS1.CancelUpdate
'rs.MoveNext
'Wend
End If
'MsgBox "three step"
If id = 0 Or id = 1 Then
    prdtoinv
End If
'MsgBox "PRDTOINV"
If id = 0 Or id = 1 Then
    'If RS1.State = 1 Then RS1.Close
    If rs2.State = 1 Then rs2.Close
    If rs3.State = 1 Then rs3.Close
    'If rs.State = 1 Then rs.Close
    'CN1.Execute "delete from pnv"
    'RS1.Open "select * from pnv", CN1, adOpenDynamic, adlockpessimistic
    'rs.Open "select * from inv", CN1, adOpenDynamic, adlockpessimistic
    rs2.Open "select prd_idy,prd_nme,grp_idy from prd", CN1, adOpenDynamic, adLockPessimistic
    rs3.Open "select grp_idy,grp_cat,grp_cls from pgr", CN1, adOpenDynamic, adLockPessimistic
    'CN1.Execute "delete from pnv"
    invtopnv
'    rs.MoveFirst
'    Set tb = dbs.TableDefs(4)
'    While Not rs.EOF
'    RS1.AddNew
'    For i = 0 To rs.Fields.Count - 1
'    fld = rs.Fields(i).name
'    RS1(fld) = rs(fld)
'    Next
'    RS1.Update
'    RS1.MoveNext
'    rs.MoveNext
'    Wend
    While Not rs2.EOF
        CN1.Execute "update pnv set prd_nme='" & Trim(rs2("prd_nme")) & "' where prd_idy like '" & rs2("prd_idy") & "'"
        CN1.Execute "update pnv set grp_idy='" & Trim(rs2("grp_idy")) & "' where prd_idy like '" & rs2("prd_idy") & "'"
        rs2.MoveNext
    Wend
    While Not rs3.EOF
        CN1.Execute "update pnv set grp_cat='" & Trim(rs3("grp_cat")) & "' where grp_idy like '" & rs3("grp_idy") & "'"
        CN1.Execute "update pnv set grp_cls='" & Trim(rs3("grp_cls")) & "' where grp_idy like '" & rs3("grp_idy") & "'"
        rs3.MoveNext
    Wend
    If rs3.State = 1 Then rs3.Close
    rs3.Open "select * from brd", CN1, adOpenDynamic, adLockPessimistic
    While Not rs3.EOF
        CN1.Execute "update pnv set bat_nme='" & Trim(rs3("bat_nme")) & "' where bat_idy like '" & rs3("bat_idy") & "'"
        CN1.Execute "update pnv set mrp_rat=" & Trim(rs3("prd_mrp")) & " where bat_idy like '" & rs3("bat_idy") & "'"
        rs3.MoveNext
    Wend
    'pb.Value = pb.Value + 5
End If
'MsgBox "fourth step"
If id = 0 Or id = 1 Then
If rs2.State = 1 Then rs2.Close
If rs3.State = 1 Then rs3.Close
If RS1.State = 1 Then RS1.Close
rs3.Open "select * from inl", CN1, adOpenDynamic, adLockPessimistic
RS1.Open "select * from pnl", CN1, adOpenDynamic, adLockPessimistic
rs2.Open "select * from act", CN1, adOpenKeyset, adLockPessimistic
'CN1.Execute "delete from pnl"
'Set tb = dbs.TableDefs(3)
'While Not rs3.EOF
'RS1.AddNew
'For i = 0 To rs3.Fields.Count - 1
'fld = rs3.Fields(i).name
'RS1(fld) = rs3(fld)
'Next
'RS1.Update
'rs3.MoveNext
'Wend
'MsgBox "completed"
I = 0
While Not rs2.EOF
    I = I + 1
    rs2.MoveNext
Wend
'MsgBox "OK"
If rs2.RecordCount > 0 Then
    rs2.MoveFirst
    For n = 0 To I - 1
        pr = rs2("act_nme")
        pr1 = chk(pr)
        CN1.Execute "update pnl set cst_nme='" & Trim(pr1) & "' where cst_idy like '" & rs2("act_idy") & "'"
        CN1.Execute "update pnl set cst_ad1='" & Trim(rs2("act_ad1")) & "' where cst_idy like '" & rs2("act_idy") & "'"
        CN1.Execute "update pnl set cst_ad2='" & Trim(rs2("act_ad2")) & "' where cst_idy like '" & rs2("act_idy") & "'"
        CN1.Execute "update pnl set cst_ad3='" & Trim(rs2("act_ad3")) & "' where cst_idy like '" & rs2("act_idy") & "'"
        CN1.Execute "update pnl set cst_ad4='" & Trim(rs2("act_ad4")) & "' where cst_idy like '" & rs2("act_idy") & "'"
        CN1.Execute "update pnl set cst_cst='" & Trim(rs2("act_cst")) & "' where cst_idy like '" & rs2("act_idy") & "'"
        CN1.Execute "update pnl set cst_gst='" & Trim(rs2("act_gst")) & "' where cst_idy like '" & rs2("act_idy") & "'"
        CN1.Execute "update pnl set cst_zip='" & Trim(rs2("act_zip")) & "' where cst_idy like '" & rs2("act_idy") & "'"
        CN1.Execute "update pnl set loc_idy='" & Trim(rs2("loc_idy")) & "' where cst_idy like '" & rs2("act_idy") & "'"
        CN1.Execute "update pnl set cat_idy='" & Trim(rs2("cat_idy")) & "' where cst_idy like '" & rs2("act_idy") & "'"
        CN1.Execute "update pnl set mar_idy='" & Trim(rs2("mar_idy")) & "' where cst_idy like '" & rs2("act_idy") & "'"
        rs2.MoveNext
    Next
End If
'MsgBox "ITS OK"
If RS1.State = 1 Then RS1.Close
RS1.Open "select * from dir where dir_idy='SYS' and dir_idx='TT1'", CN1, adOpenKeyset, adLockPessimistic
If RS1.RecordCount > 0 Then
    CN1.Execute "UPDATE pnl set ins_nme='" & RS1("dir_msg") & "'"
End If
If RS1.State = 1 Then RS1.Close
RS1.Open "select * from dir where dir_idy='SYS' and dir_idx='TT2'", CN1, adOpenKeyset, adLockPessimistic
If RS1.RecordCount > 0 Then
    CN1.Execute "UPDATE pnl set ins_adr='" & RS1("dir_msg") & "'"
End If
If RS1.State = 1 Then RS1.Close
RS1.Open "select * from dir where dir_idy='SYS' and dir_idx='GST'", CN1, adOpenKeyset, adLockPessimistic
If RS1.RecordCount > 0 Then
    CN1.Execute "UPDATE pnl set ins_gst='" & RS1("dir_msg") & "'"
End If
If RS1.State = 1 Then RS1.Close
RS1.Open "select * from dir where dir_idy='SYS' and dir_idx='CST'", CN1, adOpenKeyset, adLockPessimistic
If RS1.RecordCount > 0 Then
    CN1.Execute "UPDATE pnl set ins_cst='" & RS1("dir_msg") & "'"
End If
If RS1.State = 1 Then RS1.Close
RS1.Open "select mar_idy,mar_nme from mar", CN1, adOpenKeyset, adLockPessimistic
While Not RS1.EOF
    CN1.Execute "UPDATE pnl set mar_nme='" & RS1("mar_nme") & "' where mar_idy like '" & Trim(RS1("mar_idy")) & "'"
    RS1.MoveNext
Wend
If RS1.State = 1 Then RS1.Close
RS1.Open "select smn_idy,smn_nme from smn", CN1, adOpenDynamic, adLockPessimistic
While Not RS1.EOF
    CN1.Execute "UPDATE pnl set smn_nme='" & RS1("smn_nme") & "' where smn_idy like '" & Trim(RS1("smn_idy")) & "'"
    RS1.MoveNext
Wend
If RS1.State = 1 Then RS1.Close
RS1.Open "select loc_idy,loc_nme from loc", CN1, adOpenDynamic, adLockPessimistic
While Not RS1.EOF
    CN1.Execute "UPDATE pnl set loc_nme='" & RS1("loc_nme") & "' where loc_idy like '" & Trim(RS1("loc_idy")) & "'"
    RS1.MoveNext
Wend
If RS1.State = 1 Then RS1.Close
RS1.Open "select cat_idy,cat_nme from cat", CN1, adOpenDynamic, adLockPessimistic
While Not RS1.EOF
    CN1.Execute "UPDATE pnl set cat_nme='" & RS1("cat_nme") & "' where cat_idy like '" & Trim(RS1("cat_idy")) & "'"
    RS1.MoveNext
Wend
'MsgBox "completed"
End If
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
If id = 0 Or id = 1 Then
    If All.Value = 0 Then
        rs.Open "SELECT * FROM IDY", cn, adOpenKeyset, adLockPessimistic
            For I = 0 To list1.ListCount - 1
                If Trim(list1.List(I)) <> "csr" And Trim(list1.List(I)) <> "ast" Then
                    FLD = "idy_" & Trim(list1.List(I))
                    CN1.Execute "UPDATE IDY SET IDY_" & Trim(list1.List(I)) & " = '" & rs(Trim(FLD)) & "'"
                End If
            Next
            CN1.Execute "UPDATE IDY SET IDY_BAT=IDY_PRD"
    End If
End If
'MsgBox "OK"
If rs.State = 1 Then rs.Close
If RS1.State = 1 Then RS1.Close
If id = 2 Then
    If All.Value = 0 Then
        rs.Open "SELECT * FROM IDY", cn, adOpenKeyset, adLockPessimistic
        While Not rs.EOF
            For I = 0 To list1.ListCount - 1
                If Trim(list1.List(I)) = "ast" Then
                    FLD = "idy_asm"
                    CN1.Execute "UPDATE IDY SET IDY_" & Trim(list1.List(I)) & " = '" & rs(Trim(FLD)) & "'"
                ElseIf Trim(list1.List(I)) = "brd" Then
                    FLD = "idy_bat"
                    CN1.Execute "UPDATE IDY SET IDY_BAT = '" & rs(Trim(FLD)) & "'"
                Else
                    FLD = "idy_" & Trim(list1.List(I))
                    CN1.Execute "UPDATE IDY SET IDY_" & Trim(list1.List(I)) & " = '" & rs(Trim(FLD)) & "'"
                End If
            Next
            
            rs.MoveNext
        Wend
    End If
End If
'MsgBox "OK"
'For k1 = pb.Value To pb.Max
'pb.Value = k1
'Next
Label3.Caption = "Completed"
MsgBox "Data Import Completed Successfully", vbExclamation, "DAS Version 4.0"
Me.MousePointer = vbDefault
Unload Me
Else
' for Efflinx
    MsgBox "Efflinx Import"
    IMPFROMEFFX
End If
Exit Sub
'errh:
'If Err.Number = 3265 Then
'Resume Next
'ElseIf Err.Number = -2147217900 Then
'Resume Next
'Else
'MsgBox Err.Description & "   " & Err.Number
'End If
errh:
If err.Number = 3265 Then
    Resume Next
ElseIf err.Number = -2147217900 Then
    Resume Next
ElseIf err.Number = -2147217887 Then
'sb = tb.name & "  " & fld
'    MsgBox err.Description
'MsgBox "3rd "
    If RS1.Fields(FLD).Type = adVarChar Or RS1.Fields(FLD).Type = adChar Or RS1.Fields(FLD).Type = adBSTR Or RS1.Fields(FLD).Type = adLongVarChar Or RS1.Fields(FLD).Type = adLongVarWChar Then
        RS1(FLD) = ""
        Resume Next
    ElseIf RS1.Fields(FLD).Type = adInteger Or RS1.Fields(FLD).Type = adNumeric Or RS1.Fields(FLD).Type = adBigInt Or RS1.Fields(FLD).Type = adSingle Or RS1.Fields(FLD).Type = adDouble Or RS1.Fields(FLD).Type = adBigInt Or RS1.Fields(FLD).Type = adCurrency Or RS1.Fields(FLD).Type = adNumeric Or RS1.Fields(FLD).Type = adSmallInt Or RS1.Fields(FLD).Type = adDecimal Then
        RS1(FLD) = 0
        Resume Next
    ElseIf RS1.Fields(FLD).Type = adDate Or RS1.Fields(FLD).Type = adDBDate Or RS1.Fields(FLD).Type = adDBTimeStamp Then
        RS1(FLD) = Date
        Resume Next
    Else
        MsgBox err.Description & "  " & err.Number, vbCritical, "DAS Version 4.0"
    End If
ElseIf err.Number = -2147467259 Then
'MsgBox "4th"
        fnm = Trim(Mid(err.Description, 44, 8))
        'If fnm = "ble 'Q11" Then
        'MsgBox Err.Description & " " & Err.Number & "  " & tb.name & "  " & fld
        'End If
        'MsgBox fnm
        found = 0
        For P = 1 To Len(Trim(fnm))
        If Mid(Trim(fnm), P, 1) = "_" Then
        found = 1
        Exit For
        End If
        Next
        If found = 1 Then
        If RS1(fnm).Type = adVarChar Or RS1(fnm).Type = adBSTR Or RS1(fnm).Type = adChar Or RS1(fnm).Type = adLongVarChar Or RS1(fnm).Type = adSmallInt Then
                RS1(fnm) = ""
                Resume
        ElseIf RS1(fnm).Type = adInteger Or RS1(fnm).Type = adNumeric Or RS1(fnm).Type = adBigInt Or RS1(fnm).Type = adSingle Or RS1(fnm).Type = adDouble Then
                RS1(fnm) = 0
                Resume
        ElseIf RS1(fnm).Type = adDate Or RS1(fnm).Type = adDBDate Or RS1(fnm).Type = adDBTimeStamp Then
                RS1(fnm) = Date
                Resume
        Else
                MsgBox err.Description & " " & err.Number & fnm
        End If
        End If
        'MsgBox Err.Description & "  " & Err.Number & "   " & Mid(Err.Description, 44, 8)
Else
'MsgBox "5th"
        MsgBox err.Description & "  " & err.Number & FLD & " " & id 'tb.name &
End If
'errh:
'If Err.Number = 3265 Then
'    Resume Next
'ElseIf Err.Number = -2147217900 Then
'Resume Next
'ElseIf Err.Number = -2147217887 Then
'If RS1(fld).Type = adChar Or RS1(fld).Type = adVarChar Or RS1(fld).Type = adBSTR Or RS1(fld).Type = adLongVarChar Then
'RS1(fld) = ""
'Resume Next
'ElseIf RS1(fld).Type = adNumeric Or RS1(fld).Type = adBigInt Or RS1(fld).Type = adInteger Or RS1(fld).Type = adSingle Or RS1(fld).Type = adDouble Then
'RS1(fld) = 0
'Resume Next
'ElseIf RS1(fld).Type = adDate Or RS1(fld).Type = adDBDate Or RS1(fld).Type = adDBTimeStamp Then
'RS1(fld) = Date
'Resume Next
'End If
'ElseIf Err.Number = -2147467259 Then
''fields does not exist in inl and not null
'fnm = Trim(Mid(Err.Description, 45, 7))
'found = 0
'For K = 1 To Len(fnm)
'If Mid(Trim(fnm), K, 1) = "_" Then
'found = 1
'Exit For
'End If
'Next
'If found = 1 Then
'If RS1(fnm).Type = adChar Or RS1(fnm).Type = adVarChar Or RS1(fnm).Type = adBSTR Or RS1(fnm).Type = adLongVarChar Then
''cn1.Execute "update pnl set " & fnm & "='.'"
'RS1(fnm) = ""
'Resume Next
'ElseIf RS1(fnm).Type = adNumeric Or RS1(fnm).Type = adBigInt Or RS1(fnm).Type = adInteger Or RS1(fnm).Type = adSingle Or RS1(fld).Type = adDouble Then
''cn1.Execute "update pnl set " & fnm & "= 0"
'RS1(fnm) = 0
'Resume Next
'ElseIf RS1(fnm).Type = adDate Or RS1(fnm).Type = adDBDate Or RS1(fnm).Type = adDBTimeStamp Then
''cn1.Execute "update pnl set " & fnm & "=" & Date
'RS1(fnm) = Date
'Resume Next
'End If
'End If
'Else
'MsgBox Err.Description & " " & Err.Number & "  " & fld
'End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub CSR_CHK_Click()
'If opt_efx.Value = True Then
'If CHECKTABLE("cSR") = 0 Then
'MsgBox "Table does not exist!", vbOKOnly + vbInformation, "Data Import"
'CSR_CHK.Value = 0
'Else
'CSR_CHK.Value = 1
'End If
'End If
If CSR_CHK.Value = 1 Then
list1.AddItem "csr"
Else
For I = 0 To list1.ListCount - 1
If list1.List(I) = "csr" Then
list1.RemoveItem I
Exit For
End If
Next
End If
End Sub

Private Sub Dir1_Change()
On Error Resume Next
st = Dir1.Path
If Mid(Dir1.Path, Len(Dir1.Path), 1) = "\" Then
st = Mid(Dir1.Path, 1, Len(Dir1.Path) - 1)
End If

If Dir(st & "\eff.dbc") = "" Then
    lblvalidfolder = UCase("Invalid DAS Version 4.0 Data folder")
Else
    lblvalidfolder = UCase("Valid DAS Version 4.0 Data folder")
End If

'Set RST = CNX.OpenSchema(adSchemaTables)
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
list1.CLEAR
With LIST2
.AddItem "irt"
.AddItem "pnl"
.AddItem "pnv"
.AddItem "brd"
.AddItem "sch"
.AddItem "lds"
.AddItem "lnv"
End With
opt_eff.Value = True
End Sub

Private Sub Form_Unload(cancel As Integer)
If cn.State = 1 Then cn.Close
If CN1.State = 1 Then CN1.Close
If cn2.State = 1 Then cn2.Close
End Sub

Private Sub GRP_CHK_Click()
'If opt_efx.Value = True Then
'If CHECKTABLE("GRP") = 0 Then
'MsgBox "Table does not exist!", vbOKOnly + vbInformation, "Data Import"
'GRP_CHK.Value = 0
'Else
'GRP_CHK.Value = 1
'End If
'End If
If GRP_CHK.Value = 1 Then
list1.AddItem "grp"
Else
For I = 0 To list1.ListCount - 1
If list1.List(I) = "grp" Then
list1.RemoveItem I
Exit For
End If
Next
End If
End Sub

Private Sub LOC_CHK_Click()
'If opt_efx.Value = True Then
'If CHECKTABLE("LOC") = 0 Then
'MsgBox "Table does not exist!", vbOKOnly + vbInformation, "Data Import"
'LOC_CHK.Value = 0
'Else
'LOC_CHK.Value = 1
'End If
'End If
If LOC_CHK.Value = 1 Then
list1.AddItem "loc"
Else
For I = 0 To list1.ListCount - 1
If list1.List(I) = "loc" Then
list1.RemoveItem I
Exit For
End If
Next
End If
End Sub

Private Sub MAR_CHK_Click()
'If opt_efx.Value = True Then
'If CHECKTABLE("MAR") = 0 Then
'MsgBox "Table does not exist!", vbOKOnly + vbInformation, "Data Import"
'MAR_CHK.Value = 0
'Else
'MAR_CHK.Value = 1
'End If
'End If
If mar_chk.Value = 1 Then
list1.AddItem "mar"
Else
For I = 0 To list1.ListCount - 1
If list1.List(I) = "mar" Then
list1.RemoveItem I
Exit For
End If
Next
End If
End Sub

Private Sub opt_eff_Click()
All.Enabled = True
list1.CLEAR
'MsgBox App.Path
End Sub

Private Sub opt_efx_Click()
All.Enabled = False
list1.CLEAR
'MsgBox App.Path
'strattributes = "DBQ=" & Dir1.Path & vbCr & "Description=Effmcg Back Up" & vbCr & "SourceType=DBF" & vbCr & "Deleted=Yes"
'DBEngine.RegisterDatabase "effx", "Microsoft Visual Foxpro Driver", True, strattributes
'CNX.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=EFFx"
'If CNX.State = 1 Then CNX.Close
'CNX.Open "effx"
'strattributes = "DBQ=" & App.Path & "\Eff.dbc" & vbCr & "Description=Effmcg Back Up" & vbCr & "SourceType=DBC" & vbCr & "Deleted=Yes"
'DBEngine.RegisterDatabase "EFFg", "Microsoft Visual Foxpro Driver", True, strattributes
'CNG.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=EFFg"
'If CNG.State = 1 Then CNG.Close
'CNG.Open "EFFg"
End Sub

Private Sub PGR_CHK_Click()
'If opt_efx.Value = True Then
'If CHECKTABLE("PGR") = 0 Then
'MsgBox "Table does not exist!", vbOKOnly + vbInformation, "Data Import"
'PGR_CHK.Value = 0
'Else
'PGR_CHK.Value = 1
'End If
'End If
If PGR_CHK.Value = 1 Then
list1.AddItem "pgr"
Else
For I = 0 To list1.ListCount - 1
If list1.List(I) = "pgr" Then
list1.RemoveItem I
Exit For
End If
Next
End If
End Sub

Private Sub PRD_CHK_Click()
'If opt_efx.Value = True Then
'If CHECKTABLE("PRD") = 0 Then
'MsgBox "Table does not exist!", vbOKOnly + vbInformation, "Data Import"
'PRD_CHK.Value = 0
'Else
'PRD_CHK.Value = 1
'End If
'End If
If PRD_CHK.Value = 1 Then
    list1.AddItem "prd"
    list1.AddItem "brd"
Else
    For I = 0 To list1.ListCount - 1
        If list1.List(I) = "prd" Then
        list1.RemoveItem I
        Exit For
        End If
    Next
    For I = 0 To list1.ListCount - 1
        If list1.List(I) = "brd" Then
            list1.RemoveItem I
            Exit For
        End If
    Next
End If
End Sub

Private Sub SCH_CHK_Click()
'If opt_efx.Value = True Then
'If CHECKTABLE("SCH") = 0 Then
'MsgBox "Table does not exist!", vbOKOnly + vbInformation, "Data Import"
'SCH_CHK.Value = 0
'Else
'SCH_CHK.Value = 1
'End If
'End If
If SCH_CHK.Value = 1 Then
list1.AddItem "SCH"
Else
For I = 0 To list1.ListCount - 1
If list1.List(I) = "SCH" Then
list1.RemoveItem I
Exit For
End If
Next
End If
End Sub

Private Sub SHL_CHK_Click()
'If opt_efx.Value = True Then
'If CHECKTABLE("SHL") = 0 Then
'SHL_CHK.Value = 0
'MsgBox "Table does not exist!", vbOKOnly + vbInformation, "Data Import"
'Else
'SHL_CHK.Value = 1
'End If
'End If
If SHL_CHK.Value = 1 Then
list1.AddItem "SHL"
Else
For I = 0 To list1.ListCount - 1
If list1.List(I) = "SHL" Then
list1.RemoveItem I
Exit For
End If
Next
End If
End Sub
Private Sub SMN_CHK_Click()
'If opt_efx.Value = True Then
'If CHECKTABLE("SMN") = 0 Then
'MsgBox "Table does not exist!", vbOKOnly + vbInformation, "Data Import"
'SMN_CHK.Value = 0
'Else
'SMN_CHK.Value = 1
'End If
'End If
If smn_chk.Value = 1 Then
list1.AddItem "smn"
Else
For I = 0 To list1.ListCount - 1
If list1.List(I) = "smn" Then
list1.RemoveItem I
Exit For
End If
Next
End If
End Sub
Public Function chk(ss As String)
Dim k2 As Integer
k2 = 0
For k1 = 1 To Len(Trim(ss))
If Mid(Trim(ss), k1, 1) = "'" Then
k2 = 1
Exit For
End If
P = P & Mid(Trim(ss), k1, 1)
Next
If k2 = 1 Then
For K = k1 + 1 To Len(Trim(ss))
P = P & Mid(Trim(ss), K, 1)
Next
Else
P = Trim(ss)
End If
chk = P
End Function
Public Sub IMPFROMEFFX()
'ON ERROR GoTo ERRH5
'strattributes = "DBQ=" & Dir1.Path & vbCr & "Description=Effmcg Back Up" & vbCr & "SourceType=DBf" & vbCr & "Deleted=Yes"
'DBEngine.RegisterDatabase "effx", "Microsoft Visual Foxpro Driver", True, strattributes
'CNX.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=EFFx"
'If CNX.State = 1 Then CNX.Close
'CNX.Open "effx"
'If CNG.State = 1 Then CNG.Close
'CNG.Open "EFF"
''strattributes = "DBQ=" & CURDIR & "\Eff.dbc" & vbCr & "Description=Effmcg Back Up" & vbCr & "SourceType=DBC" & vbCr & "Deleted=Yes"
''DBEngine.RegisterDatabase "EFFg", "Microsoft Visual Foxpro Driver", True, strattributes
''CNG.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=EFFg"
''If CNG.State = 1 Then CNG.Close
''CNG.Open "EFFg"
'Set RST = CNX.OpenSchema(adSchemaTables)
'Set RST1 = CNG.OpenSchema(adSchemaTables)
''While Not RST.EOF
' '   MsgBox RST!TABLE_NAME
'  '  RST.MoveNext
''Wend
'While Not RST1.EOF
'List3.AddItem RST1!TABLE_NAME
'RST1.MoveNext
'Wend
'If rid.State = 1 Then rid.Close
'rid.Open "select * from idy", CNG, adOpenKeyset, adlockpessimistic
''MsgBox list1.ListCount - 1
'For i1 = 0 To List1.ListCount - 1
'    For n1 = 0 To List3.ListCount - 1
'       If Trim(List1.List(i1)) = Trim(List3.List(n1)) Then
'        If RSX.State = 1 Then RSX.Close
'        If RSG.State = 1 Then RSG.Close
'        RSX.Open "SELECT * FROM " & Trim(List1.List(i1)), CNX, adOpenKeyset, adlockpessimistic
'        RSG.Open "SELECT * FROM " & Trim(List1.List(i)), CNG, adOpenKeyset, adlockpessimistic
'        While Not RSX.EOF
'            RSG.AddNew
'            For FF = 0 To RSG.Fields.Count - 1
'                RSG(FF) = RSX(FF)
'            Next
'            If Trim(List1.List(i)) = "ast" Then
''                RSG("idy_asn") = Sincrement(rid("idy_asn"))
'            Else
'                fld = "idy_" & Trim(List1.List(i))
'                FLD1 = Trim(List1.List(i)) & "_IDY"
'                RSG(FLD1) = increment(rid(fld))
'            End If
'        RSG.Update
'        RSX.MoveNext
'        Wend
'           MsgBox List1.List(i)
'        Exit For
'        End If
'    Next
'Next
'MsgBox "COMPELTE"
'ERRH5:
'MsgBox err.Description & " " & err.Number & "  " & list1.List(i) & " " & RSG(FF)
End Sub

'Public Function CHECKTABLE(TB As String)
'found = 0
'RST.MoveFirst
''Dim RST As New ADODB.Recordset
'Do While Not RST.EOF
'If TB = RST!TABLE_NAME Then
'found = 1
'Exit Do
'End If
'RST.MoveNext
'Loop
'CHECKTABLE = found
'End Function
