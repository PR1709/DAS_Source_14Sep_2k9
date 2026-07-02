VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form m_FRM_HUP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAS Mobile - Retailer Information "
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   Icon            =   "m_FRM_HUP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleMode       =   0  'User
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   -15
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   585
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Retailer &Information"
      TabPicture(0)   =   "m_FRM_HUP.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LVMOD"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdsel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "MSComm1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "m_FRM_HUP.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "List1"
      Tab(1).ControlCount=   2
      Begin MSCommLib.MSComm MSComm1 
         Left            =   7410
         Top             =   5100
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.CommandButton cmdsel 
         Caption         =   "Select &All"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   6615
         Begin VB.OptionButton optall 
            Caption         =   "&4 All"
            Height          =   255
            Left            =   5520
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optboth 
            Caption         =   "&3 Both"
            Height          =   255
            Left            =   3960
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optrou 
            Caption         =   "&2 Route"
            Height          =   255
            Left            =   2280
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optsal 
            Caption         =   "&1 Salesman"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1780
         Left            =   6960
         TabIndex        =   5
         Top             =   480
         Width           =   1695
         Begin VB.CommandButton Command3 
            Caption         =   "&Send to HH"
            Height          =   395
            Left            =   90
            TabIndex        =   20
            Top             =   765
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "C&reate File"
            Height          =   395
            Left            =   120
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   285
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   395
            Left            =   105
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1245
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1150
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   6615
         Begin VB.TextBox rou_idy 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   720
            Width           =   5295
         End
         Begin VB.TextBox sal_idy 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox sal_smn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            ForeColor       =   &H80000005&
            Height          =   330
            Left            =   2400
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Route"
            Height          =   195
            Left            =   615
            TabIndex        =   4
            Top             =   720
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Salesman"
            Height          =   195
            Left            =   360
            TabIndex        =   3
            Top             =   315
            Width           =   690
         End
      End
      Begin MSComctlLib.ListView List1 
         Height          =   4500
         Left            =   -74760
         TabIndex        =   8
         Top             =   840
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   7938
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Order ID"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order Date"
            Object.Width           =   4304
         EndProperty
      End
      Begin MSComctlLib.ListView LVMOD 
         Height          =   2685
         Left            =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2415
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   4736
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Retailer Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Market"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Retailer Due"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Retailer ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Salesman"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -74775
         TabIndex        =   9
         Top             =   465
         Width           =   8055
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8730
      _ExtentX        =   15399
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
            Key             =   "n"
            Object.ToolTipText     =   "New (Ctrl+N)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ed"
            Object.ToolTipText     =   "Edit (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Area ID"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "brw"
            Object.ToolTipText     =   "Browse (Ctrl+B)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   5280
         TabIndex        =   21
         Top             =   150
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   195
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6480
         Top             =   120
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
               Picture         =   "m_FRM_HUP.frx":047A
               Key             =   "p"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "m_FRM_HUP.frx":09BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "m_FRM_HUP.frx":0D86
               Key             =   "c"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "m_FRM_HUP.frx":10DA
               Key             =   "e"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "m_FRM_HUP.frx":1712
               Key             =   "s"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "m_FRM_HUP.frx":1A66
               Key             =   "h"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "m_FRM_HUP.frx":1EC6
               Key             =   "ED"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "m_FRM_HUP.frx":21BA
               Key             =   "pr"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "m_FRM_HUP.frx":25C6
               Key             =   "d"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "m_FRM_HUP.frx":26D2
               Key             =   "pre"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "m_FRM_HUP.frx":2A26
               Key             =   "brw"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "m_FRM_HUP.frx":2E3A
               Key             =   "cpy"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "m_FRM_HUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BOTH As Boolean
Dim STR As String
Dim HHT_RST1 As New ADODB.Recordset
Dim INT_FLG As Boolean
Dim TOT_REC, rec_len, pb1, FILE_NAME
Dim I, J, Comport As Integer
Dim d1, t1, f_name, response, in_str, input_data, dumm
Public Sub delay(dlytime As Variant)
    Dim delay_timer
    Dim dli
    delay_timer = Timer
    dli = 0
    Do
    dli = dli + 1
       If Timer > (delay_timer + dlytime) Then
          Exit Do
       End If
    Loop
End Sub

Private Sub Command1_KeyPress(K As Integer)
If K = 27 Then
Unload Me
End If
End Sub

Private Sub Command2_KeyPress(K As Integer)
If K = 27 Then
Unload Me
End If
End Sub
Private Sub Command3_Click()
Command3.Enabled = False
Command2.SetFocus
in_str = ""
    TOT_REC = 0
    rec_len = 0
    MSComm1.InBufferCount = 0
    Call settings
    FILE_NAME = "RET.TXT"
    FILE_NAME = CURDIR & "\HHT" & "\" & FILE_NAME
    If FILE_NAME = "" Then
    Exit Sub
    End If
    Open FILE_NAME For Input As #1
    If EOF(1) Then
        MsgBox "Retailers File Not Available or File Creation Error!", vbCritical, "DAS Version 4.0"
        Exit Sub
    End If
    Do
        Line Input #1, input_data
        TOT_REC = TOT_REC + 1
        rec_len = Len(input_data)
    Loop While Not EOF(1)
    Close #1
      
    Do
        dumm = DoEvents()
        If MSComm1.PortOpen = True Then 'UPDATE TO CHECK THE PORT STATUS BY PKB
        MSComm1.Output = "$"
        COM_ERR
        delay (1)
        If MSComm1.PortOpen = True Then 'UPDATE TO CHECK THE PORT STATUS BY PKB
        If MSComm1.Input = "@" Then Exit Do
        Else                'FOR ABOVE IF CONDITION
        Exit Sub            'FOR ABOVE IF CONDITION
        End If              'FOR ABOVE IF CONDITION
        End If
    Loop

    For I = 1 To 4
        If Len(TOT_REC) < I Then
            TOT_REC = "0" & TOT_REC
        End If
    Next I
    
    For I = 1 To 4
        MSComm1.Output = Mid(TOT_REC, I, 1)
    Next I
    in_str = ""

    For I = 1 To 3
        If Len(rec_len) < I Then
            rec_len = "0" & rec_len
        End If
    Next I
    For I = 1 To 3
        MSComm1.Output = Mid(rec_len, I, 1)
    Next I
    delay (1)
    f_name = Dir(FILE_NAME)
    For I = 1 To 10
        If I <= Len(f_name) Then
            MSComm1.Output = Mid(f_name, I, 1)
        Else
            MSComm1.Output = " "
        End If
    Next I
    Do
        in_str = MSComm1.Input
        Loop While in_str <> "#"
    Open FILE_NAME For Input As #1
    For I = 1 To TOT_REC
        in_str = ""
        Line Input #1, input_data
        For J = 1 To rec_len
            If J <= Len(input_data) Then
                MSComm1.Output = Mid(input_data, J, 1)
            Else
                MSComm1.Output = " "
            End If
        Next J
        MSComm1.Output = " "
        Do
        in_str = MSComm1.Input
        Loop While in_str <> "#"
    Next I
    Close #1
    MsgBox "Retailers Information Send Successful to Hand Held!", vbInformation, "DAS Version 4.0"
    Command3.Enabled = True
    End Sub
Private Sub Form_KeyDown(K As Integer, S As Integer)
If K = 120 Then
    If osversion = 1 Then Shell "c:\windows\calc.exe"
    If osversion = 2 Then Shell "c:\winnt\system32\calc.exe"
End If
    If K = 78 And S = 2 Then
        newrec
        sal_idy.SetFocus
    ElseIf K = 83 And S = 2 Then
        save
    ElseIf K = 81 And S = 2 Then
        cancel
    ElseIf K = 88 And S = 2 Then
        exitfrm
    ElseIf K = 67 And S = 2 Then
    ElseIf K = 68 And S = 2 Then
        delete
    ElseIf K = 69 And S = 2 Then
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And S = 2 Then
    ElseIf K = 82 And S = 2 Then
End If
End Sub
Private Sub Form_Load()
Command1.Enabled = False
Command3.Enabled = False
list1.ListItems.CLEAR
    INT_FLG = False
End Sub
Private Sub settings()
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    If Option1.Value = True Then
        Comport = 1
    Else
        Comport = 2
    End If
    With MSComm1
        .CommPort = Comport
        .InBufferSize = 3072
        .OutBufferSize = 512
        .InBufferCount = 0
        .InputLen = 1
        .settings = "57600,n,8,1"
    End With
    If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
End Sub
Private Sub cmdsel_Click()
If cmdsel.Caption = "Select &All" Then
    For I = 1 To LVMOD.ListItems.COUNT
        LVMOD.ListItems(I).Checked = True
        Command1.Enabled = True
    Next
    cmdsel.Caption = "&Deselect All"
Else
    For I = 1 To LVMOD.ListItems.COUNT
        LVMOD.ListItems(I).Checked = False
    Next
    Command1.Enabled = False
    cmdsel.Caption = "Select &All"
End If
End Sub
Private Sub Command1_Click()
If MsgBox("Create Retailer File for Hand Held Export?", vbYesNo, "DAS Version 4.0") = vbYes Then
    Dim f
    Dim FS_SOU As Object
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    If FS_SOU.FolderExists(CURDIR & "\HHT") = False Then
    FS_SOU.CreateFolder (CURDIR & "\HHT")
    m_FRM_HUP.Visible = False
    End If
    If FS_SOU.FileExists(CURDIR & "\HHT\RET.TXT") = True Then
    FS_SOU.DeleteFile (CURDIR & "\HHT\RET.TXT")
    End If
    Set f = FS_SOU.OpenTextFile(CURDIR & "\HHT\RET.TXT", 2, True)
    Dim J As Integer
    J = 0
    For I = 1 To LVMOD.ListItems.COUNT
    If LVMOD.ListItems(I).Checked = True Then
        Dim DIS_VALUE, LEN_DISV, BAL_SPACE As Integer
         DIS_VALUE = LVMOD.ListItems(I).SubItems(2)
'         DIS_VALUE = Trim(DIS_VALUE)
        LEN_DISV = Len(DIS_VALUE)
        BAL_SPACE = 10 - LEN_DISV
        DIS_VALUE = DIS_VALUE & Space(BAL_SPACE)
        f.WriteLine LVMOD.ListItems(I).SubItems(3) & Mid(LVMOD.ListItems(I).Text, 1, 30) & DIS_VALUE
        J = J + 1
    Else
        LVMOD.Refresh
    End If
   Next I
   If J > 0 Then
      MsgBox "" & J & ": Retailers Information Created for Hand Held Export!"
      Command3.Enabled = True
      cancel
    Else
      MsgBox "Retailers  not selected for Hand Held Export!"
      newrec
      If (optall.Value = False) And (optrou.Value = False) And (optboth.Value = False) Then
         sal_idy.SetFocus
      Else
         optall.SetFocus
      End If
  End If
    f.Close
End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
    SSTab1.Tab = 0
    SSTab1.TabEnabled(1) = False
    Frame2.Enabled = True
    If INT_FLG = False Then
        optsal.SetFocus
        INT_FLG = True
    End If
End Sub

Private Sub List1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub List1_DblClick()
    list1_KeyDown 13, 0
End Sub
Private Sub list1_KeyDown(KeyCode As Integer, Shift As Integer)
list1.Refresh
If Label2.Caption = "Select Route" Then
    Dim rs As New ADODB.Recordset
    If KeyCode = 13 Then
        If list1.ListItems.COUNT > 0 Then
           If rs.State = 1 Then rs.Close
           Dim Lst1 As String
            Lst1 = list1.SelectedItem.Text
            rs.Open "select * from RDY where rou_idy like '" & Trim(Lst1) & "'", CON, adOpenStatic
            rou_idy.Text = Trim(rs("rou_idy"))
            list1.Visible = False
            SSTab1.Tab = 0
            '========================================
            If STR <> "" Then
                    If STR = "All" Then
                        HHT_RST1.Open "select act.act_idy,act.act_nme,mar.mar_nme from act,mar where act.mar_idy=mar.mar_idy", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Salesman" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT,mar WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy and RDY.SMN_IDY='" & sal_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Route" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT,mar WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy and RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Both" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT,mar WHERE ACT.ACT_IDY=RDY.ACT_IDY AND act.mar_idy=mar.mar_idy and RDY.SMN_IDY='" & sal_idy.Text & "' AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    End If
                    '--------------------------------DUES LIST SHOULD BE UPTED HERE----------------------
                    ACT_DUE_LST
                    '------------------DUES LIST SHOULD BE UPTED HERE-----------------------
                    If HHT_RST1.State = 1 Then HHT_RST1.Close
                    LVMOD.SetFocus
                   ' Command1.Enabled = True
                Else
                    MsgBox "Selection Should be Made!", vbInformation + vbYesNo, "DAS Version 4.0"
                    If STR = "All" Then
                        optall.SetFocus
                    ElseIf STR = "Salesman" Then
                        optsal.SetFocus
                    ElseIf STR = "Route" Then
                        optrou.SetFocus
                    ElseIf STR = "Both" Then
                        optboth.SetFocus
                    End If
                End If
        End If
    End If
Else
     If KeyCode = 13 Then
        If list1.ListItems.COUNT > 0 Then
            If rs.State = 1 Then rs.Close
            Dim lst As ListSubItem
            Set lst = list1.SelectedItem.ListSubItems.Item(1)
            rs.Open "select * from smn where smn_idy like '" & Trim(lst.Text) & "'", CON, adOpenStatic
            sal_idy.Text = Trim(rs("smn_idy"))
            sal_smn.Text = Trim(rs("smn_nme"))
            list1.Visible = False
            SSTab1.Tab = 0
            If BOTH = True Then
                LVMOD.SetFocus
            Else
                If STR <> "" Then
                    If STR = "All" Then
                        HHT_RST1.Open "select act.act_idy,act.act_nme,mar.mar_nme from act,mar where act.mar_idy=mar.mar_idy", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Salesman" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,mar.mar_nme from RDY,ACT,mar WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy and RDY.SMN_IDY='" & sal_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Route" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,mar.mar_nme from RDY,ACT,mar WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy and RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Both" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,mar.mar_nme from RDY,ACT,mar WHERE ACT.ACT_IDY=RDY.ACT_IDY AND act.mar_idy=mar.mar_idy and RDY.SMN_IDY='" & sal_idy.Text & "' AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    End If
                    ACT_DUE_LST
                    If HHT_RST1.State = 1 Then HHT_RST1.Close
                    SSTab1.TabEnabled(1) = False
                    LVMOD.SetFocus
                   ' Command1.Enabled = True
                Else
                    SSTab1.TabEnabled(1) = False
                    MsgBox "Selection Should Be Made!", vbInformation + vbYesNo, "DAS Version 4.0"
                    optsal.SetFocus
                End If
            End If
        End If
    End If
End If
End Sub
Private Sub List1_keypress(KeyAscii As Integer)
If KeyAscii = 27 Then
list1.Visible = False
SSTab1.Tab = 0
newrec
End If
End Sub

Private Sub list1_LostFocus()
If optboth.Value = True And rou_idy.Text = "" Then
rou_idy.SetFocus
Else
End If
End Sub

Private Sub LVMOD_Click()
Dim COUNT As Integer
    COUNT = 0
        If LVMOD.ListItems.COUNT > 0 Then
            For I = 1 To LVMOD.ListItems.COUNT
            If (LVMOD.ListItems(I).Checked = True) Then
            COUNT = COUNT + 1
            Else
            COUNT = COUNT
            End If
            Next
        End If
        If COUNT > 0 Then

       Command1.Enabled = True
       Else
            Command1.Enabled = False
         End If

End Sub

Private Sub LVMOD_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If LVMOD.SortOrder = lvwAscending Then
LVMOD.SortOrder = lvwDescending
ElseIf LVMOD.SortOrder = lvwDescending Then
LVMOD.SortOrder = lvwAscending
End If
LVMOD.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub LVMOD_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If LVMOD.SelectedItem.Checked = True Then
Command1.Enabled = True
Else
'Command1.Enabled = False
End If

End Sub

Private Sub LVMOD_KeyPress(K As Integer)
    Dim COUNT As Integer
    COUNT = 0
    If K = 27 Then
        If LVMOD.ListItems.COUNT > 0 Then
            For I = 1 To LVMOD.ListItems.COUNT
            If (LVMOD.ListItems(I).Checked = True) Then
            COUNT = COUNT + 1
            Else
            COUNT = COUNT
            End If
            Next
        If Not COUNT > 0 Then
            If optsal.Value = True Then
                sal_idy.SetFocus
            ElseIf optrou.Value = True Then
                rou_idy.SetFocus
            ElseIf optboth.Value = True Then
                sal_idy.SetFocus
            ElseIf optall.Value = True Then
                Command2.SetFocus
            End If
            newrec
       Else
            Command1.SetFocus
         End If
        Else
            Command2.SetFocus
            newrec
        End If
        End If
        
        
'    COUNT = 0
'    If K = 27 Then
'        If LVMOD.ListItems.COUNT > 0 Then
'        If LVMOD.ListItems(LVMOD.SelectedItem.Index).Checked = False Then
'            If optsal.Value = True Then
'                sal_idy.SetFocus
'            ElseIf optrou.Value = True Then
'                rou_idy.SetFocus
'            ElseIf optboth.Value = True Then
'                sal_idy.SetFocus
'            ElseIf optall.Value = True Then
'                Command2.SetFocus
'            End If
'            newrec
'       ElseIf LVMOD.ListItems(LVMOD.SelectedItem.Index).Checked = True Then
'       ' Command1.Enabled = True
'       Else
'            Command1.SetFocus
'         End If
'        Else
'            Command2.SetFocus
'            newrec
'        End If
'        End If
'    If K = 13 And Command3.Enabled = True Then
'        If LVMOD.ListItems(LVMOD.SelectedItem.Index).Checked = False Then
'            Command2.SetFocus
'        ElseIf LVMOD.ListItems(LVMOD.SelectedItem.Index).Checked = True Then
'            Command1.SetFocus
'        End If
'        Else
'        Command1.Enabled = False
'        cancel
'        If optrou.Value = True Then
'        optrou.SetFocus
'        ElseIf optsal.Value = True Then
'        optsal.SetFocus
'        ElseIf optboth.Value = True Then
'        optboth.SetFocus
'        ElseIf optall.Value = True Then
'        optall.SetFocus
'        Else
'        Command2.SetFocus
'        End If
'        End If
'  '  End If
End Sub

Private Sub LVMOD_LostFocus()
'Dim COUNT As Integer
'COUNT = 0
'For I = 1 To LVMOD.ListItems.COUNT
'        LVMOD.ListItems(I).Checked = True
'        COUNT = COUNT + 1
'        'Command1.Enabled = True
'    Next
'If Not COUNT > 0 Then
' MsgBox "No data selected for create file!", vbInformation, "DAS Version 4.0"
' LVMOD.SetFocus
' End If
End Sub

Private Sub optall_Click()
cancel
Command1.Visible = True
sal_idy.Visible = False
Label1.Visible = False
sal_smn.Visible = False
rou_idy.Visible = False
Label3.Visible = False
STR = "All"
End Sub

Private Sub optall_GotFocus()
Command1.Enabled = False
End Sub

Private Sub optall_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'If LVMOD.ListItems.COUNT > 0 Then
       ' Command1.Enabled = True
           ' Command1.SetFocus
       ' Else
         '   Command2.SetFocus
       ' End If
'cancel
Frame2.Enabled = False
STR = "All"
'========================
If STR <> "" Then
    If STR = "All" Then
        HHT_RST1.Open "select act.act_idy,act.act_nme,mar.mar_nme from act,mar where act.mar_idy=mar.mar_idy", CON, adOpenKeyset, adLockPessimistic
    ElseIf STR = "Salesman" Then
        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
    ElseIf STR = "Route" Then
        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
    ElseIf STR = "Both" Then
        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE ACT.ACT_IDY=RDY.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
    End If
    '---------------DUES LIST SHOULD BE UPTED HERE---------------------------
    ACT_DUE_LST
    '---------------------------DUES LIST SHOULD BE UPTED HERE---------------------------
    If HHT_RST1.State = 1 Then HHT_RST1.Close
    LVMOD.SetFocus
Else
    MsgBox "Selection Should Be Made!", vbInformation + vbYesNo, "DAS Version 4.0"
    optall.SetFocus
End If
End If
End Sub

Private Sub optall_KeyPress(K As Integer)
If K = 27 Then
Unload Me
End If
End Sub
Private Sub optboth_Click()
Command1.Enabled = False
    cancel
    Frame2.Enabled = True
    Frame2.BackColor = &H8000000F
    BOTH = True
    sal_idy.Visible = True
    Label1.Visible = True
    sal_smn.Visible = True
    rou_idy.Visible = True
    Label3.Visible = True
    STR = "Both"
    SET_POS
End Sub
Private Sub optboth_KeyDown(KeyCode As Integer, Shift As Integer)
Command1.Enabled = False
If KeyCode = 13 Then
    cancel
    Frame2.Enabled = True
    BOTH = True
    sal_idy.Visible = True
    Label1.Visible = True
    sal_smn.Visible = True
    rou_idy.Visible = True
    Label3.Visible = True
    SET_POS
    STR = "Both"
    sal_idy.SetFocus
End If
End Sub
Private Sub optboth_KeyPress(K As Integer)
If K = 27 Then
Unload Me
End If
End Sub
Private Sub optrou_Click()
    Command1.Enabled = False
    cancel
    Frame2.Enabled = True
    Frame2.BackColor = &H8000000F
    BOTH = False
    Label1.Visible = False
    sal_idy.Visible = False
    sal_smn.Visible = False
    rou_idy.Visible = True
    Label3.Visible = True
    SET_POS
    STR = "Route"
End Sub
Private Sub optrou_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cancel
        Frame2.Enabled = True
        BOTH = False
        Label1.Visible = False
        sal_idy.Visible = False
        sal_smn.Visible = False
        rou_idy.Visible = True
        Label3.Visible = True
        SET_POS
        rou_idy.SetFocus
    End If
End Sub
Private Sub optrou_KeyPress(K As Integer)
If K = 27 Then
Unload Me
End If
End Sub
Private Sub optsal_Click()
        cancel
        Frame2.Enabled = True
        Frame2.BackColor = &H8000000F
        BOTH = False
        Label3.Visible = False
        rou_idy.Visible = False
        sal_idy.Visible = True
        Label1.Visible = True
        sal_smn.Visible = True
        SET_POS
        STR = "Salesman"
End Sub

Private Sub optsal_GotFocus()
Command1.Enabled = False
End Sub

Private Sub optsal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cancel
        Frame2.Enabled = True
        Frame2.BackColor = &H8000000F
        BOTH = False
        Label3.Visible = False
        rou_idy.Visible = False
        sal_idy.Visible = True
        sal_idy.Enabled = True
        Label1.Visible = True
        sal_smn.Visible = True
        SET_POS
        STR = "Salesman"
        sal_idy.SetFocus
    End If
End Sub
Private Sub optsal_KeyPress(K As Integer)
If K = 27 Then
    Unload Me
End If
End Sub

Private Sub PBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

End Sub

Private Sub rou_idy_GotFocus()
    M_MDI.STS_BAR.Panels(1).Text = "Type Route ID or Press <Enter> To Get The List of Route ID's"
End Sub
Private Sub rou_idy_KeyDown(K As Integer, Shift As Integer)
If K = 69 And Shift = 2 Then
EDITDATA
End If
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Then
    If optsal.Value = True Then
        STR = "Salesman"
    ElseIf optrou.Value = True Then
        STR = "Route"
    ElseIf optboth.Value = True Then
        STR = "Both"
    End If
    LVMOD.ListItems.CLEAR
    If rs.State = 1 Then rs.Close
    If Len(Trim(rou_idy)) <> 0 Then
        rs.Open "select DISTINCT ROU_IDY from rdy where rou_idy like '" & Trim(rou_idy.Text) & "%'", CON, adOpenStatic
        If rs.RecordCount > 0 Then
            If rs.RecordCount = 1 Then
                rou_idy.Text = rs("ROU_IDY")
                If STR <> "" Then
                    If STR = "All" Then
                        HHT_RST1.Open "select act.act_idy,act.act_nme,mar.mar_nme from act,mar where act.mar_idy=mar.mar_idy", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Salesman" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Route" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Both" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE ACT.ACT_IDY=RDY.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    End If
                    '--------------------------DUES LIST SHOULD BE UPTED HERE-----------------------------------
                    ACT_DUE_LST
                    '--------------------DUES LIST SHOULD BE UPTED HERE-----------------------------------------
                    If HHT_RST1.State = 1 Then HHT_RST1.Close
                    cmdsel.SetFocus
                Else
                    MsgBox "Selection Should Be Made!", vbInformation + vbYesNo, "DAS Version 4.0"
                    optrou.SetFocus
                End If
            Else
                Label2.Caption = "Select Route"
                
                list1.ColumnHeaders.CLEAR
                list1.ColumnHeaders.Add , , "Route"
                list1.ListItems.CLEAR
                For A = 1 To rs.RecordCount
                    Set ls = list1.ListItems.Add(, , Trim(rs("ROU_IDY")))
                    rs.MoveNext
                Next
                SSTab1.Tab = 1
                list1.Visible = True
                list1.ListItems(1).Selected = True
                list1.SetFocus
            End If
        Else
            If SSTab1.Tab = 0 Then
            MsgBox "Route Not Found, Define Route !", vbExclamation, "DAS Version 4.0"
            SSTab1.Tab = 0
            rou_idy.Enabled = True
            rou_idy.SetFocus
            End If
        End If
    Else
    SSTab1.TabEnabled(1) = True
        rs.Open "select DISTINCT ROU_IDY from rdy", CON, adOpenStatic
        Label2.Caption = "Select Route"
        list1.ColumnHeaders.CLEAR
        list1.ColumnHeaders.Add , , "Route"
        list1.ListItems.CLEAR
        If rs.RecordCount > 0 Then
            For A = 1 To rs.RecordCount
            If Len(Trim("ROU_IDY")) > 0 Then
                Set ls = list1.ListItems.Add(, , Trim(rs("ROU_IDY")))
                End If
                rs.MoveNext
            Next
            SSTab1.Tab = 1
            list1.Visible = True
            list1.ListItems(1).Selected = True
            list1.SetFocus
            
        Else
            If SSTab1.Tab = 0 Then
            MsgBox "Route Not Found, Define Route !", vbExclamation, "DAS Version 4.0"
            SSTab1.Tab = 0
            rou_idy.Enabled = True
            rou_idy.SetFocus
            End If
        End If
    End If
    If rs.State = 1 Then rs.Close
End If
End Sub
Private Sub rou_idy_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 27 And optrou.Value = True Then
Unload Me
End If

End Sub
Private Sub sal_IDY_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Salesmen ID or Press <Enter> To Get The List of Salesmen ID's"
End Sub
Private Sub sal_idy_KeyDown(K As Integer, Shift As Integer)
If K = 69 And Shift = 2 Then
EDITDATA
End If
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Then
    If optsal.Value = True Then
        STR = "Salesman"
    ElseIf optrou.Value = True Then
        STR = "Route"
    ElseIf optboth.Value = True Then
        STR = "Both"
    End If
    LVMOD.ListItems.CLEAR
    cmdsel.SetFocus
    If rs.State = 1 Then rs.Close
    If Val(sal_idy) > 0 Then
    rs.Open "select * from smn where smn_idy like '" & Sincrement(sal_idy.Text) & "'", CON, adOpenStatic
    If rs.RecordCount > 0 Then
        sal_idy.Text = Sincrement(sal_idy)
        sal_smn.Text = Trim(rs("smn_nme"))
        If optboth.Value = True Then
        rou_idy.SetFocus
        Else
        If STR <> "" Then
                    If STR = "All" Then
                        HHT_RST1.Open "select act.act_idy,act.act_nme,mar.mar_nme from act,mar where act.mar_idy=mar.mar_idy", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Salesman" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Route" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Both" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE ACT.ACT_IDY=RDY.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    End If
                    '----------------DUES LIST SHOULD BE UPTED HERE ---------------------
                    ACT_DUE_LST
                    '----------------DUES LIST SHOULD BE UPTED HERE-----------------------
                    If HHT_RST1.State = 1 Then HHT_RST1.Close
                    cmdsel.SetFocus
                Else
                    MsgBox "Selection is not made!", vbInformation + vbYesNo, "DAS Version 4.0"
                    optsal.SetFocus
                End If
        End If
    End If
    ElseIf Trim(Len(sal_idy)) > 0 And Val(sal_idy) = 0 Then
        rs.Open "select * from smn where smn_nme like '" & Trim(sal_idy.Text) & "%'", CON, adOpenStatic
        If rs.RecordCount = 1 Then
            sal_idy.Text = Trim(rs("smn_IDY"))
            sal_smn.Text = Trim(rs("smn_nme"))
            If optboth.Value = True Then
            rou_idy.SetFocus
            Else
            If STR <> "" Then
                    
                    If STR = "All" Then
                        HHT_RST1.Open "select act.act_idy,act.act_nme,mar.mar_nme from act,mar where act.mar_idy=mar.mar_idy", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Salesman" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Route" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Both" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE ACT.ACT_IDY=RDY.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    End If
                    '-----------------
                    ACT_DUE_LST
                    If HHT_RST1.State = 1 Then HHT_RST1.Close
                Else
                    MsgBox "Selection is not made!", vbInformation + vbYesNo, "DAS Version 4.0"
                    optsal.SetFocus
                End If
                cmdsel.SetFocus
            End If
        Else
        smnlist
        End If
    Else
    smnlist
    End If
End If
If rs.State = 1 Then rs.Close
End Sub
Private Sub sal_IDY_KeyPress(K As Integer)
If K = 27 Then
Unload Me
End If
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
End Sub
Private Sub cancel()
STR = ""
LVMOD.ListItems.CLEAR
clrctr m_FRM_HUP
Command1.Enabled = False
cmdsel.Caption = "Select &All" 'CHANGED BY PKB
idnew = False
SSTab1.Tab = 0
End Sub
Private Sub smnlist()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
SSTab1.TabEnabled(1) = True
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
If Val(sal_idy.Text) = 0 And Len(Trim(sal_idy.Text)) = 0 Then
rs.Open "select * from SMN", CON, adOpenStatic
ElseIf Len(Trim(sal_idy)) <> 0 Then
rs.Open "select * from smn where smn_nme like '" & sal_idy & "%'", CON, adOpenStatic
End If
Label2.Caption = "Select Salesman"
list1.ColumnHeaders.CLEAR
list1.ColumnHeaders.Add , , "Salesman"
list1.ColumnHeaders.Add , , "ID"
list1.ColumnHeaders.Add , , "Address"
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("smn_nme")))
        ls.ListSubItems.Add , , Trim(rs("SMN_IDY") & "")
        ls.ListSubItems.Add , , Trim(rs("SMN_AD4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.ListItems(1).Selected = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        sal_idy.Enabled = True
        sal_idy.SetFocus
    End If
End If
End Sub
Public Sub SET_POS()
    If BOTH = True Then
        Label1.Top = 315
        sal_idy.Top = 240
        sal_smn.Top = 240
        Label3.Top = 750
        rou_idy.Top = 720
    ElseIf optsal.Value = True Then
        sal_idy.Top = 400
        sal_smn.Top = 400
        Label1.Top = 475
    ElseIf optrou.Value = True Then
        rou_idy.Top = 400
        Label3.Top = 430
    End If

End Sub
Public Sub ACT_DUE_LST()
Dim PP_RST1 As New ADODB.Recordset
Dim TT_RST1 As New ADODB.Recordset
Dim DUE() As Double
ReDim DUE(HHT_RST1.RecordCount) As Double
TRAIL_BALANCE Format(Date, "MM/DD/YYYY")

LVMOD.Visible = True
LVMOD.ListItems.CLEAR
Dim TRL_BAL As Double
While HHT_RST1.EOF = False
If TT_RST1.State = 1 Then TT_RST1.Close
TT_RST1.Open "SELECT * FROM OPA WHERE ACT_IDY IN( SELECT ACT_IDY FROM ACT WHERE GRP_CLF='Debtor' and act_idy='" & HHT_RST1(0) & "')", CON, adOpenKeyset
If TT_RST1.EOF = False Then
    TRL_BAL = Val(TT_RST1(1)) - Val(TT_RST1(2))
Else
    TRL_BAL = 0
End If
        Dim S As Integer
        Dim S1 As Integer
        Dim MRP_RAT1 As String
        MRP_RAT1 = convert(CStr(TRL_BAL))
        S = InStr(Trim(MRP_RAT1), ".")
        S1 = Len(Mid(Trim(MRP_RAT1), 1, S - 1))
        If 5 - S1 = 1 Then
            MRP_RAT1 = Space(1) & Trim(MRP_RAT1)
        ElseIf 5 - S1 = 2 Then
            MRP_RAT1 = Space(2) & Trim(MRP_RAT1)
        ElseIf 5 - S1 = 3 Then
            MRP_RAT1 = Space(3) & Trim(MRP_RAT1)
        ElseIf 5 - S1 = 4 Then
            MRP_RAT1 = Space(4) & Trim(MRP_RAT1)
        End If
LVMOD.GetFirstVisible
Set ls = LVMOD.ListItems.Add(, , HHT_RST1(1) & "")
ls.ListSubItems.Add , , HHT_RST1(2) & ""
ls.ListSubItems.Add , , MRP_RAT1 & ""
ls.ListSubItems.Add , , HHT_RST1(0) & ""
HHT_RST1.MoveNext
Wend
While HHT_RST1.EOF = False
    Set ls = LVMOD.ListItems.Add(, , HHT_RST1(0) & "")
    ls.ListSubItems.Add , , HHT_RST1(1) & ""
    HHT_RST1.MoveNext
Wend
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
    newrec
ElseIf Button.KEY = "s" Then
    save
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "e" Then
    exitfrm
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "Ed" And optsal.Value = True Then
EDITDATA
ElseIf Button.KEY = "Ed" And optrou.Value = True Then
If rs.State = 1 Then rs.Close
rs.Open "select DISTINCT ROU_IDY from rdy where rou_idy like '" & Trim(rou_idy.Text) & "%'", CON, adOpenStatic
        If rs.RecordCount > 0 Then
Label2.Caption = "Select Route"
                list1.ColumnHeaders.CLEAR
                list1.ColumnHeaders.Add , , "Route"
                list1.ListItems.CLEAR
                For A = 1 To rs.RecordCount
                    Set ls = list1.ListItems.Add(, , Trim(rs("ROU_IDY")))
                    rs.MoveNext
                Next
                SSTab1.Tab = 1
                list1.Visible = True
                list1.ListItems(1).Selected = True
                list1.SetFocus
ElseIf Button.KEY = "pr" Then
ElseIf Button.KEY = "crt" Then
End If
End If
End Sub
Public Sub newrec()
clrctr m_FRM_HUP
LVMOD.ListItems.CLEAR
End Sub
Public Sub save()
If MsgBox("Confirm Record Save ?", vbYesNo + vbExclamation, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
End Sub

Public Sub exitfrm()
Unload Me
End Sub
Public Sub delete()
End Sub
Public Sub EDITDATA()
Dim rs As New ADODB.Recordset
'If K = 13 Or K = 40 Then
    If optsal.Value = True Then
        STR = "Salesman"
    ElseIf optrou.Value = True Then
        STR = "Route"
    ElseIf optboth.Value = True Then
        STR = "Both"
    End If
    LVMOD.ListItems.CLEAR
    cmdsel.SetFocus
    If rs.State = 1 Then rs.Close
    If Val(sal_idy) > 0 Then
    rs.Open "select * from smn where smn_idy like '" & Sincrement(sal_idy.Text) & "'", CON, adOpenStatic
    If rs.RecordCount > 0 Then
        sal_idy.Text = Sincrement(sal_idy)
        sal_smn.Text = Trim(rs("smn_nme"))
        If optboth.Value = True Then
        rou_idy.SetFocus
        Else
        If STR <> "" Then
                    If STR = "All" Then
                        HHT_RST1.Open "select act.act_idy,act.act_nme,mar.mar_nme from act,mar where act.mar_idy=mar.mar_idy", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Salesman" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Route" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Both" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE ACT.ACT_IDY=RDY.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    End If
                    '----------------DUES LIST SHOULD BE UPTED HERE ---------------------
                    ACT_DUE_LST
                    If HHT_RST1.State = 1 Then HHT_RST1.Close
                    cmdsel.SetFocus
                Else
                    MsgBox "Selection is not made!", vbInformation + vbYesNo, "DAS Version 4.0"
                    optsal.SetFocus
                End If
        End If
    End If
    ElseIf Trim(Len(sal_idy)) > 0 And Val(sal_idy) = 0 Then
        rs.Open "select * from smn where smn_nme like '" & Trim(sal_idy.Text) & "%'", CON, adOpenStatic
        If rs.RecordCount = 1 Then
            sal_idy.Text = Trim(rs("smn_IDY"))
            sal_smn.Text = Trim(rs("smn_nme"))
            If optboth.Value = True Then
            rou_idy.SetFocus
            Else
            If STR <> "" Then
                    If STR = "All" Then
                        HHT_RST1.Open "select act.act_idy,act.act_nme,mar.mar_nme from act,mar where act.mar_idy=mar.mar_idy", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Salesman" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Route" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE RDY.ACT_IDY=ACT.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    ElseIf STR = "Both" Then
                        HHT_RST1.Open "select RDY.ACT_IDY,ACT.ACT_NME,MAR.MAR_NME from RDY,ACT WHERE ACT.ACT_IDY=RDY.ACT_IDY AND act.mar_idy=mar.mar_idy AND RDY.SMN_IDY='" & sal_idy.Text & "' AND RDY.ROU_IDY='" & rou_idy.Text & "' GROUP BY RDY.ACT_IDY", CON, adOpenKeyset, adLockPessimistic
                    End If
                    '-----------------
                    ACT_DUE_LST
                    If HHT_RST1.State = 1 Then HHT_RST1.Close
                Else
                    MsgBox "Selection Should Be Made!", vbInformation + vbYesNo, "DAS Version 4.0"
                    optsal.SetFocus
                End If
                cmdsel.SetFocus
            End If
        Else
        smnlist
        End If
    Else
    smnlist
    End If
If rs.State = 1 Then rs.Close
End Sub

'Public Sub COM_ERR()
'   Select Case MSComm1.CommEvent
'      Case comEventBreak   ' A Break was received.
'      MsgBox "A COM Break Detected! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventFrame   ' Framing Error
'      MsgBox "Framing Error! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventOverrun   ' Data Lost.
'      MsgBox "Data Transmission Loss! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventRxOver   ' Receive buffer overflow.
'      MsgBox "COM Port Data Receive Buffer Overflow!, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventRxParity   ' Parity Error.
'      MsgBox "COM Port Parity Error! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventTxFull   ' Transmit buffer full.
'      MsgBox "COM Transmit Buffer Full! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEventDCB   ' Unexpected error retrieving DCB
'      MsgBox "Error Retrieving DCB, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvCD   ' Change in the CD line.
'      MsgBox "Change in CD Line, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvCTS   ' Change in the CTS line.
'      MsgBox "Change in CTS Line, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvDSR   ' Change in the DSR line.
'      MsgBox "Change in DSR Line, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvRing   ' Change in the Ring Indicator.
'      MsgBox "Change in Ring Indicator, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvReceive   ' Received RThreshold # of
'      MsgBox "Receive Error, Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvSend   ' There are SThreshold number of
'                     ' characters in the transmit
'                     ' buffer.
'      MsgBox "Data is in Transmit Buffer! Check COM Port and Related Connections!", vbCritical, "DAS Version 4.0"
'      Case comEvEOF   ' An EOF charater was found in
'                     ' the input stream
'      MsgBox "EOF in Input Stream! Check COM Connections!", vbCritical, "DAS Version 4.0"
'   End Select
'End Sub
'
'
Public Sub COM_ERR()
    Dim RESULT As Variant
    Select Case MSComm1.CommEvent
      Case comEventBreak   ' A Break was received.
      RESULT = MsgBox("A COM Break Detected! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventFrame   ' Framing Error
      RESULT = MsgBox("Framing Error! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventOverrun   ' Data Lost.
      RESULT = MsgBox("Data Transmission Loss! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventRxOver   ' Receive buffer overflow.
      RESULT = MsgBox("COM Port Data Receive Buffer Overflow! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventRxParity   ' Parity Error.
      RESULT = MsgBox("COM Port Parity Error! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventTxFull   ' Transmit buffer full.
      RESULT = MsgBox("COM Transmit Buffer Full! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEventDCB   ' Unexpected error retrieving DCB
      RESULT = MsgBox("Error Retrieving DCB! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvCD   ' Change in the CD line.
      RESULT = MsgBox("Change in CD Line! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvCTS   ' Change in the CTS line.
      RESULT = MsgBox("Change in CTS Line! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvDSR   ' Change in the DSR line.
      RESULT = MsgBox("Change in DSR Line! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvRing   ' Change in the Ring Indicator.
      RESULT = MsgBox("Change in Ring Indicator! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvReceive   ' Received RThreshold # of
      RESULT = MsgBox("Receive Error! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
      RESULT = MsgBox("Data is in Transmit Buffer! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
      Case comEvEOF   ' An EOF charater was found in
                     ' the input stream
      RESULT = MsgBox("EOF in Input Stream! Check COM Port and Related Connections!", vbRetryCancel, "DAS Version 4.0")
   End Select
   If RESULT = 4 Or RESULT = "" Then
   'Unload Me
   ElseIf RESULT = 2 Then
   Unload Me
   'M_FRM_MIBS.Visible = True
   End If
End Sub

