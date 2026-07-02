VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_RPT_GTH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Growth Rate Report"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "M_RPT_GTH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9022
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Report Control Filters"
      TabPicture(0)   =   "M_RPT_GTH.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fdate"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "tdate"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Check8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbCaldates"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check7"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmbsku"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "&View"
      TabPicture(1)   =   "M_RPT_GTH.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "list1"
      Tab(1).Control(1)=   "Label20"
      Tab(1).ControlCount=   2
      Begin VB.ComboBox cmbsku 
         Height          =   315
         ItemData        =   "M_RPT_GTH.frx":047A
         Left            =   1800
         List            =   "M_RPT_GTH.frx":0490
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   7200
         TabIndex        =   11
         Top             =   3600
         Width           =   1335
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   240
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.CheckBox Check7 
         Caption         =   "&All"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1400
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   3375
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   5895
         Begin VB.OptionButton optRet 
            Caption         =   "&Retailer"
            Height          =   255
            Left            =   360
            TabIndex        =   36
            Top             =   2813
            Width           =   1095
         End
         Begin VB.OptionButton optRou 
            Caption         =   "Rout&e"
            Height          =   255
            Left            =   360
            TabIndex        =   35
            Top             =   2310
            Width           =   1095
         End
         Begin VB.OptionButton optSmn 
            Caption         =   "&Salesman"
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   1853
            Width           =   1095
         End
         Begin VB.OptionButton optMar 
            Caption         =   "&Market"
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   1373
            Width           =   1095
         End
         Begin VB.OptionButton optCat 
            Caption         =   "&Category"
            Height          =   375
            Left            =   360
            TabIndex        =   32
            Top             =   833
            Width           =   1095
         End
         Begin VB.OptionButton optLoc 
            Caption         =   "&Location"
            Height          =   375
            Left            =   360
            TabIndex        =   31
            Top             =   338
            Width           =   1095
         End
         Begin VB.CheckBox chkARet 
            Caption         =   "All Reta&ilers"
            Height          =   375
            Left            =   4440
            TabIndex        =   30
            Top             =   2760
            Width           =   1335
         End
         Begin VB.CheckBox chkARou 
            Caption         =   "All R&outes"
            Height          =   375
            Left            =   4440
            TabIndex        =   29
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CheckBox chkASmn 
            Caption         =   "All Salesma&n"
            Height          =   375
            Left            =   4440
            TabIndex        =   28
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CheckBox chkAMar 
            Caption         =   "All Mar&kets"
            Height          =   375
            Left            =   4440
            TabIndex        =   27
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkACat 
            Caption         =   "All Cate&gories"
            Height          =   375
            Left            =   4440
            TabIndex        =   26
            Top             =   840
            Width           =   1335
         End
         Begin VB.CheckBox chkALoc 
            Caption         =   "All Loca&tions"
            Height          =   375
            Left            =   4440
            TabIndex        =   25
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtLoc 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox txtCat 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   855
            Width           =   2655
         End
         Begin VB.TextBox txtMar 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1335
            Width           =   2655
         End
         Begin VB.TextBox txtSmn 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1815
            Width           =   2655
         End
         Begin VB.TextBox txtRet 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2775
            Width           =   2655
         End
         Begin VB.ComboBox txtRou 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2280
            Width           =   735
         End
      End
      Begin VB.ComboBox cmbCaldates 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   525
         Width           =   615
      End
      Begin VB.CheckBox Check8 
         Caption         =   "&FMCG Calendar"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   555
         Width           =   1455
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74520
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1065
         Visible         =   0   'False
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
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
            Text            =   "Market Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker tdate 
         Height          =   330
         Left            =   4920
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   525
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24707075
         CurrentDate     =   37258
      End
      Begin MSComCtl2.DTPicker fdate 
         Height          =   330
         Left            =   3240
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   525
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24707075
         CurrentDate     =   37258
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Select "
         Height          =   195
         Left            =   960
         TabIndex        =   39
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   24
         Top             =   3840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   23
         Top             =   2880
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   22
         Top             =   2400
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   21
         Top             =   1920
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3720
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   4560
         TabIndex        =   19
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   2760
         TabIndex        =   18
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Market"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   -74520
         TabIndex        =   17
         Top             =   720
         Width           =   7785
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   0
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
            Picture         =   "M_RPT_GTH.frx":04CE
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_GTH.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_GTH.frx":0DDA
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_GTH.frx":112E
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_GTH.frx":1766
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_GTH.frx":1ABA
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_GTH.frx":1F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_GTH.frx":220E
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_GTH.frx":261A
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_GTH.frx":2726
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_GTH.frx":2A7A
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_RPT_GTH.frx":2E8E
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "n"
            Object.ToolTipText     =   "New (Ctrl + N)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s"
            Object.ToolTipText     =   "Save (Ctrl + S)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "c"
            Object.ToolTipText     =   "Cancel (Ctrl+Q)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sku"
                  Object.Tag             =   "sku"
                  Text            =   "SKU Wise Route Sales"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "e"
            Object.ToolTipText     =   "Exit (Ctrl+X)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_RPT_GTH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rswek As New ADODB.Recordset

Private Sub Check7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check7.Value = 1 Then
        Frame1.Enabled = False
        Command1.SetFocus
    Else
        Frame1.Enabled = True
        optLoc.SetFocus
    End If
End If
End Sub

Private Sub cmbCaldates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If rswek.State = 1 Then rswek.Close
rswek.Open "select frm_dat,to_dat from wek where wek_idy=" & Val(cmbCaldates.Text), CON, adOpenKeyset, adLockOptimistic
If Not rswek.EOF Then
    fdate.Enabled = True
    tdate.Enabled = True
    fdate.Value = Format(rswek("frm_dat"), "MM/DD/YYYY")
    tdate.Value = Format(rswek("to_dat"), "MM/DD/YYYY")
    fdate.Enabled = False
    tdate.Enabled = False
    cmbsku.SetFocus
End If
End If
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    
End If
End Sub

Private Sub cmbsku_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Len(Trim(cmbsku.Text)) > 0 Then
        Check7.Enabled = True
        Check7.SetFocus
    Else
        MsgBox "Select an Item", vbExclamation, "Effmcg"
        cmbsku.SetFocus
    End If
End If
End Sub

Private Sub fdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    tdate.SetFocus
End If
End Sub

Private Sub Form_Activate()
cmbsku.ListIndex = 0
Check8.SetFocus
End Sub
Private Sub Form_KeyDown(K As Integer, S As Integer)
If K = 120 Then Shell "c:\windows\calc.exe"
If K > 111 And K < 120 Then If TypeOf Me.ActiveControl Is TextBox Then Me.ActiveControl.Text = Trim(FUNKEY(K))
If K = 78 And S = 2 Then
    'NewRecord
ElseIf K = 83 And S = 2 Then
    'saverecord
ElseIf K = 81 And S = 2 Then
    cancel
ElseIf K = 88 And S = 2 Then
    Unload Me
ElseIf K = 112 Then
    M_HELP
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If Me.ActiveControl.name = "Check8" Then
        Unload Me
    Else
        cancel
'        saverecord
    End If
End If
End Sub

Private Sub Form_Load()
SETYEAR Me
End Sub

Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub

Private Sub optCat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtCat.Enabled = True
    txtCat.SetFocus
End If
End Sub

Private Sub optLoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtLoc.Enabled = True
    txtLoc.SetFocus
End If
End Sub

Private Sub optMar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtMar.Enabled = True
    txtMar.SetFocus
End If
End Sub

Private Sub optRet_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtRet.Enabled = True
    txtRet.SetFocus
End If
End Sub

Private Sub optRou_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtRou.Enabled = True
    txtRou.SetFocus
End If
End Sub

Private Sub optSmn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtSmn.Enabled = True
    txtSmn.SetFocus
End If
End Sub

Private Sub tdate_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If tdate.Value >= fdate.Value Then
        cmbsku.SetFocus
    Else
        MsgBox "Invalid Date", vbExclamation, "Effmcg"
        tdate.SetFocus
    End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "n" Then
'    NewRecord
ElseIf Button.KEY = "s" Then
'    saverecord
ElseIf Button.KEY = "c" Then
 '   cancel
ElseIf Button.KEY = "p" Then

ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "e" Then
    Unload Me
End If
End Sub
Private Sub Check8_Click()
Check8_KeyDown 13, 0
End Sub

Private Sub Check8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check8.Value = 1 Then
        cmbCaldates.Enabled = True
        If rswek.State = 1 Then rswek.Close
        rswek.Open "select wek_idy from wek", CON, adOpenKeyset, adLockOptimistic
        cmbCaldates.Clear
        If Not rswek.EOF Then
            While Not rswek.EOF
                cmbCaldates.AddItem rswek("wek_idy")
                rswek.MoveNext
            Wend
            cmbCaldates.SetFocus
        Else
            MsgBox "Sales Calendar does not exist!", vbExclamation, "Effmcg"
            Check8.SetFocus
            Check8.Value = 0
        End If
    Else
        cmbCaldates.Enabled = False
        fdate.Enabled = True
        tdate.Enabled = True
        fdate.SetFocus
    End If
End If
End Sub

Private Sub txtCat_Change()
If txtCat = "" Then Label4 = ""
If Len(Trim(txtCat)) > 0 Then
    txtLoc = ""
    txtMar = ""
    txtSmn = ""
    txtRet = ""
    chkALoc.Value = 0
    chkACat.Value = 0
    chkAMar.Value = 0
    chkASmn.Value = 0
    chkARou.Value = 0
    chkARet.Value = 0
End If
End Sub

Private Sub txtLoc_Change()
If txtLoc = "" Then Label3 = ""
If Len(Trim(txtLoc)) > 0 Then
    txtCat = ""
    txtMar = ""
    txtSmn = ""
    txtRet = ""
    chkALoc.Value = 0
    chkACat.Value = 0
    chkAMar.Value = 0
    chkASmn.Value = 0
    chkARou.Value = 0
    chkARet.Value = 0
End If
End Sub

Private Sub txtloc_Keydown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("LOC", "LOC_IDY", Trim(Label3)) = True Then
Check2.SetFocus
Exit Sub
End If
RS.Open "select * from LOC where LOC_NME like '" & Trim(text1) & "%'", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("LOC_NME")))
        ls.ListSubItems.Add , , Trim(RS("LOC_IDY") & "")
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Location"
    list1.ColumnHeaders(1).Text = "Location Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Location Not Found, Define Location !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        txtLoc.Enabled = True
        txtLoc.SetFocus
    End If
End If
End If
End Sub

Private Sub txtloc_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub


Private Sub txtCat_Keydown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("CAT", "CAT_IDY", Trim(Label4)) = True Then
CHECK3.SetFocus
Exit Sub
End If

RS.Open "select * from CAT where cAT_nme like '" & Trim(Text2) & "%'", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("CAT_nme")))
        ls.ListSubItems.Add , , Trim(RS("cAT_idy") & "")
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Category"
    list1.ColumnHeaders(1).Text = "Category Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Category Not Found, Define Category...!", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        txtCat.Enabled = True
        txtCat.SetFocus
    End If
End If


End If
End Sub

Private Sub txtcat_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub Txtmar_Change()
If Text3 = "" Then Label5 = ""
If Len(Trim(txtMar)) > 0 Then
    txtCat = ""
    txtLoc = ""
    txtSmn = ""
    txtRet = ""
    chkALoc.Value = 0
    chkACat.Value = 0
    chkAMar.Value = 0
    chkASmn.Value = 0
    chkARou.Value = 0
    chkARet.Value = 0
End If
End Sub

Private Sub Txtmar_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("MAR", "MAR_IDY", Trim(Label5)) = True Then
Check4.SetFocus
Exit Sub
End If

RS.Open "select * from mar where mar_nme like '" & Trim(Text3) & "%'", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("mar_nme")))
        ls.ListSubItems.Add , , Trim(RS("mar_idy") & "")
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Market"
    list1.ColumnHeaders(1).Text = "Market Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Market Not Found, Define Market !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        txtMar.Enabled = True
        txtMar.SetFocus
    End If
End If
End If
End Sub

Private Sub txtmar_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub

Private Sub txtRou_Change()
If Len(Trim(txtRou.Text)) > 0 Then
    txtCat = ""
    txtMar = ""
    txtSmn = ""
    txtRet = ""
    txtLoc = ""
    chkALoc.Value = 0
    chkACat.Value = 0
    chkAMar.Value = 0
    chkASmn.Value = 0
    chkARou.Value = 0
    chkARet.Value = 0
End If
End Sub

Private Sub Txtsmn_Change()
If txtSmn = "" Then Label6 = ""
If Len(Trim(txtSmn)) > 0 Then
    txtCat = ""
    txtMar = ""
    txtLoc = ""
    txtRet = ""
    chkALoc.Value = 0
    chkACat.Value = 0
    chkAMar.Value = 0
    chkASmn.Value = 0
    chkARou.Value = 0
    chkARet.Value = 0
End If
End Sub

Private Sub Txtsmn_KeyDown(K As Integer, Shift As Integer)
If K = 13 Then
Dim ls As ListItem
Dim RS As New ADODB.Recordset
If check("SMN", "SMN_IDY", Trim(Label6)) = True Then
    Check5.SetFocus
    Exit Sub
End If
RS.Open "select * from smn where smn_nme like '" & Trim(Text4) & "%'", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("smn_nme")))
        ls.ListSubItems.Add , , Trim(RS("smn_idy") & "")
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Salesman"
    list1.ColumnHeaders(1).Text = "Salesman Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        txtSmn.Enabled = True
        txtSmn.SetFocus
    End If
End If
End If
End Sub

Private Sub Txtsmn_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub
Private Sub cancel()
Check8.Value = 0
cmbCaldates.Clear
fdate.Value = Date
tdate.Value = Date
cmbsku.ListIndex = 0
Check7.Value = 0
txtLoc = ""
txtCat = ""
txtMar = ""
txtSmn = ""
txtRet = ""
chkALoc.Value = 0
chkACat.Value = 0
chkAMar.Value = 0
chkASmn.Value = 0
chkARou.Value = 0
chkARet.Value = 0
optLoc.Value = False
optCat.Value = False
optMar.Value = False
optSmn.Value = False
optRou.Value = False
optRet.Value = False
Check8.SetFocus
End Sub

Private Sub Txtret_Change()
If txtRet = "" Then Label7 = ""
If Len(Trim(txtRet)) > 0 Then
    txtCat = ""
    txtMar = ""
    txtSmn = ""
    txtLoc = ""
    chkALoc.Value = 0
    chkACat.Value = 0
    chkAMar.Value = 0
    chkASmn.Value = 0
    chkARou.Value = 0
    chkARet.Value = 0
End If
End Sub

Private Sub Txtret_keydown(K As Integer, Shift As Integer)
If K = 13 Then
If check("ACT", "ACT_IDY", Trim(Label7)) = True Then
Check8.SetFocus
Exit Sub
End If
Dim ls As ListItem
Dim RS As New ADODB.Recordset
RS.Open "select * from act where act_nme like '" & Trim(Text6) & "%' and (act_typ like 'SAL' or act_typ like 'BOT')", CON, adOpenStatic
list1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(RS("act_nme")))
        ls.ListSubItems.Add , , Trim(RS("act_idy") & "")
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    Label20 = "Select Retailer"
    list1.ColumnHeaders(1).Text = "Retailer Name"
    list1.ColumnHeaders(2).Text = "ID"
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Retailer Not Found, Define Retailer !", vbExclamation, "Effmcg"
        SSTab1.Tab = 0
        txtRet.Enabled = True
        txtRet.SetFocus
    End If
End If
End If
End Sub

Private Sub Txtret_KeyPress(K As Integer)
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K > 96 And K < 123 Then K = K - 32
End Sub
Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_Validate(cancel As Boolean)
cancel = True
End Sub

Private Sub list1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If list1.ListItems.Count > 0 Then
        If Label20 = "Select Location" Then
            Label3 = list1.SelectedItem.ListSubItems.Item(1)
            txtLoc = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Category" Then
            Label4 = list1.SelectedItem.ListSubItems.Item(1)
            txtCat = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Market" Then
            Label5 = list1.SelectedItem.ListSubItems.Item(1)
            txtMar = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Salesman" Then
            Label6 = list1.SelectedItem.ListSubItems.Item(1)
            txtSmn = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        ElseIf Label20 = "Select Retailer" Then
            Label7 = list1.SelectedItem.ListSubItems.Item(1)
            txtRet = Trim(list1.SelectedItem)
            SSTab1.Tab = 0
            list1.Visible = False
            Command1.SetFocus
        End If
    End If
End If
End Sub
