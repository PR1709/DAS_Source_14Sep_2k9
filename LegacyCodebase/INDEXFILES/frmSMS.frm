VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSM Mobile Interface - Business Short Messaging Service - SMS"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   Icon            =   "frmSMS.frx":0000
   KeyPreview      =   -1  'True
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
      Top             =   4680
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
      TabPicture(0)   =   "frmSMS.frx":08CA
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
      TabPicture(1)   =   "frmSMS.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line2"
      Tab(1).Control(1)=   "rtbMsgTemplate"
      Tab(1).Control(2)=   "cmdSaveTemplate"
      Tab(1).Control(3)=   "cmdClear"
      Tab(1).Control(4)=   "lstFields"
      Tab(1).Control(5)=   "lstMsgTemplate"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "&2 Message Send"
      TabPicture(2)   =   "frmSMS.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvSMSSend"
      Tab(2).Control(1)=   "FrameMaster"
      Tab(2).Control(2)=   "Command1"
      Tab(2).Control(3)=   "lstSMSSend"
      Tab(2).Control(4)=   "Check1"
      Tab(2).Control(5)=   "FrameAccount"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "&3 Message Receive"
      TabPicture(3)   =   "frmSMS.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvReceivedMessages"
      Tab(3).Control(1)=   "Command2"
      Tab(3).Control(2)=   "Command4"
      Tab(3).Control(3)=   "cmdInterpret"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "&4 Message Log"
      TabPicture(4)   =   "frmSMS.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label6"
      Tab(4).Control(1)=   "RichTextBox1"
      Tab(4).Control(2)=   "cmdClearLog"
      Tab(4).Control(3)=   "cmdSaveLog"
      Tab(4).Control(4)=   "Command3"
      Tab(4).Control(5)=   "cmbSMSMode"
      Tab(4).Control(6)=   "cmdOpenLog"
      Tab(4).ControlCount=   7
      Begin VB.CommandButton cmdInterpret 
         Caption         =   "&InterPret"
         Height          =   375
         Left            =   -71160
         TabIndex        =   46
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Frame FrameAccount 
         Height          =   1455
         Left            =   -71760
         TabIndex        =   29
         Top             =   480
         Width           =   4455
         Begin VB.ListBox List2 
            Height          =   1035
            Left            =   1440
            MultiSelect     =   2  'Extended
            TabIndex        =   33
            Top             =   240
            Width           =   2895
         End
         Begin VB.OptionButton OptAccount 
            Caption         =   "Supplier"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   32
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton OptAccount 
            Caption         =   "Retailer"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptAccount 
            Caption         =   "Sales Man"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   30
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Se&lect All"
         Height          =   375
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Restore Defaults"
         Height          =   375
         Left            =   6480
         TabIndex        =   27
         Top             =   3960
         Width           =   1335
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
         ItemData        =   "frmSMS.frx":0956
         Left            =   -73920
         List            =   "frmSMS.frx":0963
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
         ItemData        =   "frmSMS.frx":0985
         Left            =   3525
         List            =   "frmSMS.frx":09B3
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1290
         Width           =   1575
      End
      Begin VB.ComboBox cmbParity 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSMS.frx":0A10
         Left            =   3525
         List            =   "frmSMS.frx":0A23
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1770
         Width           =   1575
      End
      Begin VB.ComboBox cmbDataBit 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSMS.frx":0A45
         Left            =   3525
         List            =   "frmSMS.frx":0A58
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2250
         Width           =   1575
      End
      Begin VB.ComboBox cmbStopBit 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmSMS.frx":0A6B
         Left            =   3525
         List            =   "frmSMS.frx":0A78
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
            ItemData        =   "frmSMS.frx":0A87
            Left            =   720
            List            =   "frmSMS.frx":0AA0
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
         TabIndex        =   34
         Top             =   2415
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   2355
         _Version        =   393217
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmSMS.frx":0AE4
      End
      Begin MSComctlLib.ListView lvSMSSend 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   35
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
         TabIndex        =   36
         Top             =   555
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5530
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sl.No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Message"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Phone No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Mode"
            Object.Width           =   2540
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2715
         Left            =   -74880
         TabIndex        =   37
         Top             =   1095
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4789
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         ScrollBars      =   3
         MaxLength       =   6000
         Appearance      =   0
         TextRTF         =   $"frmSMS.frx":0B66
      End
      Begin VB.Label Label7 
         Caption         =   "secs"
         Height          =   255
         Left            =   4080
         TabIndex        =   45
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMS Mode"
         Height          =   195
         Left            =   -74880
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   900
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baud Rate "
         Height          =   195
         Left            =   2430
         TabIndex        =   42
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parity"
         Height          =   195
         Left            =   2850
         TabIndex        =   41
         Top             =   1830
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Bit"
         Height          =   195
         Left            =   2670
         TabIndex        =   40
         Top             =   2310
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stop Bit"
         Height          =   195
         Left            =   2685
         TabIndex        =   39
         Top             =   2790
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out Setting"
         Height          =   195
         Left            =   2055
         TabIndex        =   38
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
For I = 1 To lvSMSSend.ListItems.COUNT
    lvSMSSend.ListItems(I).Checked = b
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
    For I = 0 To cmbBaudRate.ListCount
        If Trim(cmbBaudRate.List(I)) = "9600" Then
        cmbBaudRate.ListIndex = I
        Exit For
        End If
    Next
    For I = 0 To cmbParity.ListCount
        If Mid(cmbParity.List(I), 1, 1) = "N" Then
        cmbParity.ListIndex = I
        Exit For
        End If
    Next
    For I = 0 To cmbDataBit.ListCount
        If cmbDataBit.List(I) = "8" Then
        cmbDataBit.ListIndex = I
        Exit For
        End If
    Next
    For I = 0 To cmbStopBit.ListCount
        If cmbStopBit.List(I) = "1" Then
        cmbStopBit.ListIndex = I
        Exit For
        End If
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strText() As String
    cmbFilter.ListIndex = 0
    strText = CommSettings()
    For I = 0 To cmbBaudRate.ListCount
        If Trim(cmbBaudRate.List(I)) = strText(0) Then
        cmbBaudRate.ListIndex = I
        Exit For
        End If
    Next
    For I = 0 To cmbParity.ListCount
        If Mid(cmbParity.List(I), 1, 1) = strText(1) Then
        cmbParity.ListIndex = I
        Exit For
        End If
    Next
    For I = 0 To cmbDataBit.ListCount
        If cmbDataBit.List(I) = strText(2) Then
        cmbDataBit.ListIndex = I
        Exit For
        End If
    Next
    For I = 0 To cmbStopBit.ListCount
        If cmbStopBit.List(I) = strText(3) Then
        cmbStopBit.ListIndex = I
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
    For I = 0 To CNT - 1
        template = Split(ArrayTemplate(I), "|", -1, vbTextCompare)
        a1(I) = template(0)
        a2(I) = template(1)
        a3(I) = template(2)
        a4(I) = template(3)
        a5(I) = template(4)
    Next
    lstMsgTemplate.CLEAR
    lstSMSSend.CLEAR
    For I = 0 To CNT - 1
        lstMsgTemplate.AddItem a1(I)
        lstSMSSend.AddItem a1(I)
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
    MsgBox "Error While Reading" & ERRMSG, vbInformation, "DAS Version 4.0"
    End If
    RichTextBox1.Text = ""
    RichTextBox1.Text = strText
    If strText = "" Then
        MsgBox "Either Log Details are erased/New Log File", vbInformation, "DAS Version 4.0"
    End If
End Sub


Private Sub cmdSaveLog_Click()
    Dim ERRMSG As Long
    Dim strFile As String
    ERRMSG = modInnComm.WriteLogFile(RichTextBox1.Text, cmbSMSMode.ListIndex, strFile)
    If ERRMSG <> 0 Then
        MsgBox "Error While Saving" & ERRMSG, vbInformation, "DAS Version 4.0"
    Else
        MsgBox "Log Details Saved - " & strFile, vbInformation, "DAS Version 4.0"
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
    s = CStr(modInnComm.ReceiveMessage(CInt(Trim(txtCommPortNo.Text)), Me.MSComm1, 4, CInt(txtWait.Text), RichTextBox1))
    SplitReceivedMessage RichTextBox1.Text
    cmdSaveLog_Click
End Sub

Public Function SplitReceivedMessage(strReceivedMessage As String)
    Dim STR() As String
    Dim s() As String
    lvReceivedMessages.ListItems.CLEAR
    STR = Split(strReceivedMessage, vbCrLf)
    For I = 0 To UBound(STR) - 1
        STR(I) = Replace(STR(I), vbCrLf, "")
        STR(I) = Replace(STR(I), vbCr, "")
        STR(I) = Trim(STR(I))
        If Len(STR(I)) > 1 Then
            s = Split(STR(I), ",")
            Set Item = lvReceivedMessages.ListItems.Add(, , s(0))
            Item.ListSubItems.Add , , s(6)
            Item.ListSubItems.Add , , s(2)
            Item.ListSubItems.Add , , s(4)
            Item.ListSubItems.Add , , s(5)
            Item.ListSubItems.Add , , s(1)
        End If
    Next
End Function

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
        MsgBox "Template not Created", vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
    Dim ERRMSG As Long
    ERRMSG = modMessageTemplate.SaveTemplate(Trim(lstMsgTemplate.Text), Trim(rtbMsgTemplate.Text))
    If ERRMSG <> 0 Then
        MsgBox "Error While Saving" & ERRMSG, vbInformation, "DAS Version 4.0"
    Else
        MsgBox "Template Saved", vbInformation, "DAS Version 4.0"
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
    For I = 0 To UBound(t1) - 1
        lstFields.AddItem t2(I) & "                                                                   |" & t1(I)
    Next
    Dim strMessageTemplate As String
    modMessageTemplate.OpenTemplate Trim(lstMsgTemplate.Text), strMessageTemplate
    rtbMsgTemplate.Text = strMessageTemplate
End Sub

'SMS Send

Private Sub Command1_Click()
    Dim strMessageTemplate As String
    Dim STR() As String
    Dim FSO As New FileSystemObject
    Dim f As File
    Dim ts As TextStream
    Dim strMsg As String
    
    If lstSMSSend.ListIndex < 0 Then
        MsgBox "Please Select The Message Type", vbInformation, "DAS Version 4.0"
        Exit Sub
    End If
    modMessageTemplate.OpenTemplate Trim(lstSMSSend.Text), strMessageTemplate
    If FSO.FileExists(SMSFolder & "\SendMessage.txt") = False Then
        Set ts = FSO.CreateTextFile(SMSFolder & "\SendMessage.txt")
        ts.Close
    End If
    Set f = FSO.GetFile(SMSFolder & "\SendMessage.txt")
    Set ts = f.OpenAsTextStream(ForAppending)
    For I = 1 To lvSMSSend.ListItems.COUNT
        If lvSMSSend.ListItems(I).Checked = True Then
            strMsg = strMessageTemplate
            strMsg = Replace(strMsg, "#" & 0 & "#", lvSMSSend.ListItems(I).Text)
            For J = 1 To lvSMSSend.ColumnHeaders.COUNT - 1
                strMsg = Replace(strMsg, "#" & J & "#", lvSMSSend.ListItems(I).ListSubItems(J).Text)
            Next
            If a5(lstSMSSend.ListIndex) = "1" Then
                For K = 0 To list2.ListCount - 1
                     If list2.Selected(K) = True Then
                        STR = Split(list2.List(K), "|")
                        ts.Write Trim(txtCommPortNo.Text) & "|" & Trim(STR(2)) & "|" & strMsg & "|" & "5" & vbCrLf
                    End If
                Next
            Else
                ts.Write Trim(txtCommPortNo.Text) & "|" & lvSMSSend.ListItems(I).ListSubItems(lvSMSSend.ColumnHeaders.COUNT - 1) & "|" & strMsg & "|" & "5" & vbCrLf
            End If
        End If
    Next
    ts.Close
    If Dir(App.Path & "/SMSSend.exe", vbDirectory) <> "" Then
        Shell App.Path & "/SMSSend.exe"
    MsgBox "Sending Messages is done parallely in the background", vbInformation, "DAS Version 4.0"
    End If
End Sub

Private Sub lstSMSSend_Click()
    lvSMSSend.ListItems.CLEAR
    lvSMSSend.ColumnHeaders.CLEAR
    If a5(lstSMSSend.ListIndex) = "1" Then
        FrameAccount.Visible = True
        FrameMaster.Visible = False
        list2.CLEAR
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
If list2.Visible = True Then
    modSMSSend.GetMasterRecords Index, list2
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

