VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmSMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSM Mobile Interface - Business Short Messaging Service - SMS"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   Icon            =   "frmSMS.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      NullDiscard     =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   706
      TabCaption(0)   =   "&Port Settings"
      TabPicture(0)   =   "frmSMS.frx":0742
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line1(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtWait"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmbStopBit"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbDataBit"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmbParity"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbBaudRate"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdSave"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCommPortNo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command5"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "&1 Message Templates"
      TabPicture(1)   =   "frmSMS.frx":075E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line2"
      Tab(1).Control(1)=   "rtbMsgTemplate"
      Tab(1).Control(2)=   "cmdSaveTemplate"
      Tab(1).Control(3)=   "cmdClear"
      Tab(1).Control(4)=   "lstFields"
      Tab(1).Control(5)=   "lstMsgTemplate"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "&2 Message Send"
      TabPicture(2)   =   "frmSMS.frx":077A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAccount"
      Tab(2).Control(1)=   "Check1"
      Tab(2).Control(2)=   "lstSMSSend"
      Tab(2).Control(3)=   "Command1"
      Tab(2).Control(4)=   "FrameMaster"
      Tab(2).Control(5)=   "lvSMSSend"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "&3 Message Receive"
      TabPicture(3)   =   "frmSMS.frx":0796
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command4"
      Tab(3).Control(1)=   "Command2"
      Tab(3).Control(2)=   "lvReceivedMessages"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "&4 Message Log"
      TabPicture(4)   =   "frmSMS.frx":07B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdInterpret"
      Tab(4).Control(1)=   "cmdOpenLog"
      Tab(4).Control(2)=   "cmbSMSMode"
      Tab(4).Control(3)=   "Command3"
      Tab(4).Control(4)=   "cmdSaveLog"
      Tab(4).Control(5)=   "cmdClearLog"
      Tab(4).Control(6)=   "RichTextBox1"
      Tab(4).Control(7)=   "Label6"
      Tab(4).ControlCount=   8
      Begin VB.Frame FrameAccount 
         Height          =   1455
         Left            =   -71760
         TabIndex        =   30
         Top             =   480
         Width           =   4455
         Begin VB.ListBox List2 
            Height          =   1035
            Left            =   1440
            MultiSelect     =   2  'Extended
            TabIndex        =   34
            Top             =   240
            Width           =   2895
         End
         Begin VB.OptionButton OptAccount 
            Caption         =   "Supplier"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   33
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton OptAccount 
            Caption         =   "Retailer"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptAccount 
            Caption         =   "Sales Man"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   31
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Se&lect All"
         Height          =   375
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Restore Defaults"
         Height          =   375
         Left            =   6480
         TabIndex        =   28
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdInterpret 
         Caption         =   "&InterPret"
         Height          =   375
         Left            =   -70680
         TabIndex        =   27
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Xrf to DB"
         Height          =   375
         Left            =   -68520
         TabIndex        =   26
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdOpenLog 
         Caption         =   "&Get"
         Height          =   375
         Left            =   -72120
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbSMSMode 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSMS.frx":07CE
         Left            =   -73920
         List            =   "frmSMS.frx":07DB
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   630
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Erase SIM"
         Height          =   375
         Left            =   -74880
         TabIndex        =   23
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveLog 
         Caption         =   "&7 Save Log"
         Height          =   375
         Left            =   -69480
         TabIndex        =   22
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdClearLog 
         Caption         =   "&8 Refresh"
         Height          =   375
         Left            =   -68280
         TabIndex        =   21
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Receive"
         Height          =   375
         Left            =   -69960
         TabIndex        =   20
         Top             =   3960
         Width           =   1335
      End
      Begin VB.ListBox lstSMSSend 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   -74880
         TabIndex        =   19
         Top             =   510
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Send"
         Height          =   375
         Left            =   -68520
         TabIndex        =   18
         Top             =   3960
         Width           =   1335
      End
      Begin VB.ListBox lstMsgTemplate 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   -74880
         TabIndex        =   17
         Top             =   555
         Width           =   4815
      End
      Begin VB.ListBox lstFields 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   -69960
         TabIndex        =   16
         Top             =   555
         Width           =   2775
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Data Clear"
         Height          =   375
         Left            =   -69960
         TabIndex        =   15
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveTemplate 
         Caption         =   "Save &Template"
         Height          =   375
         Left            =   -68520
         TabIndex        =   14
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtCommPortNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3525
         TabIndex        =   13
         Text            =   "6"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Apply"
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   3960
         Width           =   1335
      End
      Begin VB.ComboBox cmbBaudRate 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSMS.frx":07FD
         Left            =   3525
         List            =   "frmSMS.frx":082B
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1290
         Width           =   1575
      End
      Begin VB.ComboBox cmbParity 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSMS.frx":0888
         Left            =   3525
         List            =   "frmSMS.frx":089B
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1770
         Width           =   1575
      End
      Begin VB.ComboBox cmbDataBit 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSMS.frx":08BD
         Left            =   3525
         List            =   "frmSMS.frx":08D0
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2250
         Width           =   1575
      End
      Begin VB.ComboBox cmbStopBit 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSMS.frx":08E3
         Left            =   3525
         List            =   "frmSMS.frx":08F0
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2730
         Width           =   1575
      End
      Begin VB.TextBox txtWait 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3525
         TabIndex        =   7
         Text            =   "5"
         Top             =   3210
         Width           =   495
      End
      Begin VB.Frame FrameMaster 
         Height          =   1455
         Left            =   -71760
         TabIndex        =   2
         Top             =   480
         Width           =   4455
         Begin VB.CommandButton cmdFilter 
            Caption         =   "&Filter"
            Height          =   375
            Left            =   3000
            TabIndex        =   5
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtFilterText 
            Height          =   375
            Left            =   720
            TabIndex        =   4
            Top             =   600
            Width           =   2175
         End
         Begin VB.ComboBox cmbFilter 
            Height          =   315
            ItemData        =   "frmSMS.frx":08FF
            Left            =   720
            List            =   "frmSMS.frx":0918
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   300
            Width           =   330
         End
      End
      Begin RichTextLib.RichTextBox rtbMsgTemplate 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   35
         Top             =   2415
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   2355
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmSMS.frx":095C
      End
      Begin MSComctlLib.ListView lvSMSSend 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   36
         Top             =   1995
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvReceivedMessages 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   37
         Top             =   555
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5530
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sl.No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2715
         Left            =   -74880
         TabIndex        =   38
         Top             =   1095
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4789
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmSMS.frx":09DE
      End
      Begin VB.Label Label7 
         Caption         =   "secs"
         Height          =   255
         Left            =   4080
         TabIndex        =   46
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMS Mode"
         Height          =   195
         Left            =   -74880
         TabIndex        =   45
         Top             =   690
         Width           =   795
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   -75000
         X2              =   -67065
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   0
         X2              =   8040
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Port Number"
         Height          =   195
         Left            =   1920
         TabIndex        =   44
         Top             =   900
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baud Rate "
         Height          =   195
         Left            =   2430
         TabIndex        =   43
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parity"
         Height          =   195
         Left            =   2850
         TabIndex        =   42
         Top             =   1830
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Bit"
         Height          =   195
         Left            =   2670
         TabIndex        =   41
         Top             =   2310
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stop Bit"
         Height          =   195
         Left            =   2685
         TabIndex        =   40
         Top             =   2790
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out Setting"
         Height          =   195
         Left            =   2055
         TabIndex        =   39
         Top             =   3270
         Width           =   1185
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   0
         X2              =   8040
         Y1              =   3600
         Y2              =   3600
      End
   End
End
Attribute VB_Name = "frmSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a1() As String
Dim a2() As String
Dim a3() As String
Dim a4() As String
Dim a5() As String
Dim ArrayTemplate() As String
Dim ErrResult As Long

Private Sub Check1_Click()
Dim b As Boolean
If Check1.Value = 1 Then
    b = True
Else
    b = False
End If
For i = 1 To lvSMSSend.ListItems.count
    lvSMSSend.ListItems(i).Checked = b
Next
End Sub

Private Sub cmdClear_Click()
rtbMsgTemplate.Text = ""
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFilter_Click()
lstSMSSend_Click
End Sub

Private Sub cmdSave_Click()
SaveCommSettings CLng(cmbBaudRate.Text), Mid(cmbParity.Text, 1, 1), CInt(cmbDataBit.Text), CInt(cmbStopBit.Text), CInt(txtCommPortNo.Text), CInt(txtWait.Text)
End Sub

Private Sub Command5_Click()
    txtCommPortNo.Text = "6"
    txtWait.Text = "5"
    For i = 0 To cmbBaudRate.ListCount
        If Trim(cmbBaudRate.List(i)) = "9600" Then
        cmbBaudRate.ListIndex = i
        Exit For
        End If
    Next
    For i = 0 To cmbParity.ListCount
        If Mid(cmbParity.List(i), 1, 1) = "N" Then
        cmbParity.ListIndex = i
        Exit For
        End If
    Next
    For i = 0 To cmbDataBit.ListCount
        If cmbDataBit.List(i) = "8" Then
        cmbDataBit.ListIndex = i
        Exit For
        End If
    Next
    For i = 0 To cmbStopBit.ListCount
        If cmbStopBit.List(i) = "1" Then
        cmbStopBit.ListIndex = i
        Exit For
        End If
    Next
End Sub

Private Sub Form_Load()
    Dim strText() As String
    cmbFilter.ListIndex = 0
    strText = CommSettings()
    For i = 0 To cmbBaudRate.ListCount
        If Trim(cmbBaudRate.List(i)) = strText(0) Then
        cmbBaudRate.ListIndex = i
        Exit For
        End If
    Next
    For i = 0 To cmbParity.ListCount
        If Mid(cmbParity.List(i), 1, 1) = strText(1) Then
        cmbParity.ListIndex = i
        Exit For
        End If
    Next
    For i = 0 To cmbDataBit.ListCount
        If cmbDataBit.List(i) = strText(2) Then
        cmbDataBit.ListIndex = i
        Exit For
        End If
    Next
    For i = 0 To cmbStopBit.ListCount
        If cmbStopBit.List(i) = strText(3) Then
        cmbStopBit.ListIndex = i
        Exit For
        End If
    Next
    txtCommPortNo.Text = strText(5)
    txtWait.Text = strText(4)
    If cmbSMSMode.ListCount > 0 Then cmbSMSMode.ListIndex = 0
    ReDim Preserve ArrayTemplate(0)
    ErrResult = modMessageTemplate.MessageTemplate(ArrayTemplate)
    Dim CNT As Long
    CNT = UBound(ArrayTemplate)
    Dim template() As String
    ReDim Preserve a1(CNT)
    ReDim Preserve a2(CNT)
    ReDim Preserve a3(CNT)
    ReDim Preserve a4(CNT)
    ReDim Preserve a5(CNT)
    For i = 0 To CNT - 1
        template = Split(ArrayTemplate(i), "|", -1, vbTextCompare)
        a1(i) = template(0)
        a2(i) = template(1)
        a3(i) = template(2)
        a4(i) = template(3)
        a5(i) = template(4)
    Next
    lstMsgTemplate.CLEAR
    lstSMSSend.CLEAR
    For i = 0 To CNT - 1
        lstMsgTemplate.AddItem a1(i)
        lstSMSSend.AddItem a1(i)
    Next
     lstFields.CLEAR
     FrameAccount.Visible = False
     FrameMaster.Visible = False
End Sub

Private Sub cmdClearLog_Click()
    RichTextBox1.Text = ""
End Sub

Private Sub cmdOpenLog_Click()
    Dim strText As String
    Dim ERRMSG As Long
    ERRMSG = modInnComm.ReadLogFile(cmbSMSMode.ListIndex, strText)
    If ERRMSG <> 0 Then
    MsgBox "Error While Reading" & ERRMSG
    End If
    RichTextBox1.Text = ""
    RichTextBox1.Text = strText
    If strText = "" Then
        MsgBox "Either Log Details are erased/New Log File"
    End If
End Sub


Private Sub cmdSaveLog_Click()
    Dim ERRMSG As Long
    Dim strFile As String
    ERRMSG = modInnComm.WriteLogFile(RichTextBox1.Text, cmbSMSMode.ListIndex, strFile)
    If ERRMSG <> 0 Then
        MsgBox "Error While Saving" & ERRMSG
    Else
        MsgBox "Log Details Saved - " & strFile
    End If
End Sub

Private Sub cmdInterpret_Click()
    If RichTextBox1.Text = "" Then Exit Sub
    Dim Item As ListItem
    lvReceivedMessages.ListItems.CLEAR
End Sub

Private Sub Command2_Click()
    SSTab1.Tab = 4
    cmdInterpret.Visible = True
    cmbSMSMode.ListIndex = 1
    Dim s As String
    s = CStr(modInnComm.ReceiveMessage(6, Me.MSComm1, 4, 5, RichTextBox1))
End Sub


Private Sub Command3_Click()
    cmbSMSMode.ListIndex = 2
modInnComm.DeleteMessage 6, MSComm1, 4
End Sub




Private Sub lvSMSSend_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvSMSSend.SortKey = ColumnHeader.Index - 1 And lvSMSSend.SortOrder = lvwDescending Then
        lvSMSSend.SortKey = ColumnHeader.Index - 1
        lvSMSSend.SortOrder = lvwAscending
    ElseIf lvSMSSend.SortKey = ColumnHeader.Index - 1 And lvSMSSend.SortOrder = lvwAscending Then
        lvSMSSend.SortKey = ColumnHeader.Index - 1
        lvSMSSend.SortOrder = lvwDescending
    Else
         lvSMSSend.SortKey = ColumnHeader.Index - 1
        lvSMSSend.SortOrder = lvwAscending
    End If
    lvSMSSend.Sorted = True
    
End Sub

Private Sub MSComm1_OnComm()
   Select Case MSComm1.CommEvent
   ' Handle each event or error by placing
   ' code below each case statement

   ' Errors
      Case comEventBreak   ' A Break was received.
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   ' Data Lost.
      Case comEventRxOver   ' Receive buffer overflow.
      Case comEventRxParity   ' Parity Error.
      Case comEventTxFull   ' Transmit buffer full.
      Case comEventDCB   ' Unexpected error retrieving DCB]

   ' Events
      Case comEvCD   ' Change in the CD line.
      Case comEvCTS   ' Change in the CTS line.
      Case comEvDSR   ' Change in the DSR line.
      Case comEvRing   ' Change in the Ring Indicator.
      Case comEvReceive   ' Received RThreshold # of
                        ' chars.
      Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
      Case comEvEOF   ' An EOF charater was found in
                     ' the input stream
   End Select
End Sub

'Messsage Template
Private Sub cmdSaveTemplate_Click()
    If lstMsgTemplate.ListIndex < 0 Then
        MsgBox "Template not Created"
        Exit Sub
    End If
    Dim ERRMSG As Long
    ERRMSG = modMessageTemplate.SaveTemplate(Trim(lstMsgTemplate.Text), Trim(rtbMsgTemplate.Text))
    If ERRMSG <> 0 Then
        MsgBox "Error While Saving" & ERRMSG
    Else
        MsgBox "Template Saved"
    End If
End Sub

Private Sub lstFields_Click()
rtbMsgTemplate.Text = rtbMsgTemplate.Text & "#" & lstFields.ListIndex & "#" ' & "-" & Trim(Mid(lstFields.Text, 1, InStr(1, lstFields.Text, "|") - 1)) & "#"
End Sub

Private Sub lstMsgTemplate_Click()
Dim t1() As String
Dim t2() As String
    t1 = Split(a3(lstMsgTemplate.ListIndex), ",", -1, vbTextCompare)
    t2 = Split(a4(lstMsgTemplate.ListIndex), ",", -1, vbTextCompare)
    lstFields.CLEAR
    For i = 0 To UBound(t1) - 1
        lstFields.AddItem t2(i) & "                                                                   |" & t1(i)
    Next
    Dim strMessageTemplate As String
    modMessageTemplate.OpenTemplate Trim(lstMsgTemplate.Text), strMessageTemplate
    rtbMsgTemplate.Text = strMessageTemplate
End Sub

'SMS Send

Private Sub Command1_Click()
If lstSMSSend.ListIndex < 0 Then
    MsgBox "Please Select The Message Type"
    Exit Sub
End If
    SSTab1.Tab = 4
    cmdInterpret.Visible = False
    cmbSMSMode.ListIndex = 0
    Dim strMessageTemplate As String
    Dim str() As String
    
    modMessageTemplate.OpenTemplate Trim(lstSMSSend.Text), strMessageTemplate

    If a5(lstSMSSend.ListIndex) = "4" Then
        
    End If
    For i = 1 To lvSMSSend.ListItems.count
        If lvSMSSend.ListItems(i).Checked = True Then
            strMessageTemplate = Replace(strMessageTemplate, "#" & 0 & "#", lvSMSSend.ListItems(i).Text)
            For J = 1 To lvSMSSend.ColumnHeaders.count - 1
                strMessageTemplate = Replace(strMessageTemplate, "#" & J & "#", lvSMSSend.ListItems(i).ListSubItems(J).Text)
            Next
            If a5(lstSMSSend.ListIndex) = "4" Then
                For K = 0 To List2.ListCount - 1
                     If List2.Selected(K) = True Then
                        str = Split(List2.List(K), "|")
                        modInnComm.SendMessage CInt(txtCommPortNo.Text), Me.MSComm1, Trim(str(2)), strMessageTemplate, 5, RichTextBox1
                    End If
                Next
            Else
                modInnComm.SendMessage CInt(txtCommPortNo.Text), Me.MSComm1, lvSMSSend.ListItems(i).ListSubItems(lvSMSSend.ColumnHeaders.count - 1), strMessageTemplate, 5, RichTextBox1
            End If
        End If
    Next
End Sub

Private Sub lstSMSSend_Click()
    lvSMSSend.ListItems.CLEAR
    lvSMSSend.ColumnHeaders.CLEAR
    If a5(lstSMSSend.ListIndex) = "1" Then
        FrameAccount.Visible = True
        FrameMaster.Visible = False
        List2.CLEAR
        OptAccount(0).Value = False
        OptAccount(1).Value = False
        OptAccount(2).Value = False
        modSMSSend.FillListView lvSMSSend, a2(lstSMSSend.ListIndex), 0, cmbFilter.ListIndex, txtFilterText.Text
    Else
        FrameMaster.Visible = True
        FrameAccount.Visible = False
        modSMSSend.FillListView lvSMSSend, a2(lstSMSSend.ListIndex), 1, cmbFilter.ListIndex, txtFilterText.Text
    End If
End Sub

'Private Sub List1_Click()
'    Dim i As Integer
'    Dim Value As String
'    If Option2(0).Value = True Then i = 0
'    If Option2(1).Value = True Then i = 1
'    If Option2(2).Value = True Then i = 2
'    If Option2(3).Value = True Then i = 3
'    If Option2(4).Value = True Then i = 4
'    If i = 2 Then
'    Value = Trim(list1.Text)
'    Else
'    Value = Split(list1.Text, "|")(1)
'    End If
'    modSMSSend.FillListView lvSMSSend, a2(lstSMSSend.ListIndex), 1, i, Value
'End Sub

Private Sub OptAccount_Click(Index As Integer)
If List2.Visible = True Then
    modSMSSend.GetMasterRecords Index, List2
Else
    modSMSSend.GetMasterRecordByAcc Index, lvSMSSend
End If
End Sub

Private Sub Option2_Click(Index As Integer)
'modSMSSend.GetMasterRecords Index + 3, list1

End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
        Case 0
            txtCommPortNo.SetFocus
        Case 1
            lstMsgTemplate.SetFocus
        Case 2
            lstSMSSend.SetFocus
        Case 3
            
        Case 4
        
    End Select
End Sub

