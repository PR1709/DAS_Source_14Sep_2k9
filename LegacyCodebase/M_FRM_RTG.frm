VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_RTG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activation Targets - PrePaid Telecom Business"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8850
   Icon            =   "M_FRM_RTG.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5145
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   9075
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Activation &Targets"
      TabPicture(0)   =   "M_FRM_RTG.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MS"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtTemp"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Check1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&View"
      TabPicture(1)   =   "M_FRM_RTG.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "list1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Select Filter Combination for Retailers "
         Height          =   2415
         Left            =   720
         TabIndex        =   9
         Top             =   360
         Width           =   6015
         Begin VB.ComboBox cmbRoute 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1920
            Width           =   1815
         End
         Begin VB.TextBox txtMarket 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1140
            MaxLength       =   40
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1500
            Width           =   1095
         End
         Begin VB.TextBox txtSalesman 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1140
            MaxLength       =   40
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtLocation 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1140
            MaxLength       =   40
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   660
            Width           =   1095
         End
         Begin VB.TextBox txtCategory 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1140
            MaxLength       =   40
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Category"
            Height          =   195
            Left            =   300
            TabIndex        =   23
            Top             =   315
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Location"
            Height          =   195
            Left            =   315
            TabIndex        =   22
            Top             =   735
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Salesman"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   1155
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Market"
            Height          =   195
            Left            =   435
            TabIndex        =   20
            Top             =   1575
            Width           =   495
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Route"
            Height          =   195
            Left            =   495
            TabIndex        =   19
            Top             =   1980
            Width           =   435
         End
         Begin VB.Label lblCat 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2340
            TabIndex        =   18
            Top             =   240
            Width           =   3525
         End
         Begin VB.Label lblLoc 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2340
            TabIndex        =   17
            Top             =   660
            Width           =   3525
         End
         Begin VB.Label lblSmn 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2340
            TabIndex        =   16
            Top             =   1080
            Width           =   3525
         End
         Begin VB.Label lblMar 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   2340
            TabIndex        =   15
            Top             =   1500
            Width           =   3525
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&All"
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtTemp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   6
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Load Retailers"
         Height          =   375
         Left            =   7440
         TabIndex        =   5
         Top             =   2400
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   2205
         Left            =   15
         TabIndex        =   2
         Top             =   2895
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   3889
         _Version        =   393216
         BackColor       =   -2147483633
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
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74550
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1100
         Visible         =   0   'False
         Width           =   7905
         _ExtentX        =   13944
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
            Text            =   "Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Retailers' Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   -74550
         TabIndex        =   4
         Top             =   735
         Width           =   7905
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":047A
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":09BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":0D86
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":10DA
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":1712
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":1A66
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":1EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":21BA
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":25C6
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":26D2
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":2A26
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":2E3A
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_RTG.frx":3386
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8850
      _ExtentX        =   15610
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
            Object.Visible         =   0   'False
            Key             =   "n"
            Object.ToolTipText     =   "New (Ctrl+N)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SR"
                  Text            =   "Target Analysis Report"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "ed"
            Object.ToolTipText     =   "View (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "pre"
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
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yy"
      Format          =   16580611
      CurrentDate     =   38017
   End
End
Attribute VB_Name = "M_FRM_RTG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsdir As New ADODB.Recordset
Dim rsRtg As New ADODB.Recordset
Dim rsact As New ADODB.Recordset
Dim rsrdy As New ADODB.Recordset
Dim SDATE As String
Dim MOD1 As Boolean
Dim r As Integer

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Check1.Value = 1 Then
        Command1.Enabled = True
        Command1.SetFocus
    Else
        txtCategory.Enabled = True
        txtCategory.SetFocus
    End If
End If
End Sub

Private Sub cmbRoute_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 And Shift = 1 Then
    txtMarket.Enabled = True
    txtMarket.SetFocus
    Exit Sub
End If
If KeyCode = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Command1_Click()
If rsRtg.State = 1 Then rsRtg.Close
rsRtg.Open "select * from rtg", con, adOpenKeyset, adLockPessimistic
If rsRtg.EOF Then
    MOD1 = False
    If rsact.State = 1 Then rsact.Close
'    rsact.Open "select act_idy,act_nme,key_act from act where grp_clf='Debtor' and cst_typ='Retailer' and cat_idy like '" & Trim(txtCategory) & "%' and loc_idy like '" & Trim(txtLocation) & "%' and smn_idy like '" & Trim(txtSalesman) & "%' and mar_idy like '" & Trim(txtMarket) & "%' and rou_idy like '" & Trim(cmbRoute) & "%' order by key_act desc", CON, adOpenKeyset, adlockpessimistic
    'rsact.Open "select act_idy,act_nme,key_act from act where grp_clf='Debtor' and cst_typ='Retailer' and cat_idy like '" & Trim(txtCategory) & "%' and loc_idy like '" & Trim(txtLocation) & "%' and smn_idy in(select smn_idy from rdy where smn_idy like '" & Trim(txtSalesman) & "%') and mar_idy like '" & Trim(txtMarket) & "%' and rou_idy like '" & Trim(cmbRoute) & "%' order by key_act desc", CON, adOpenKeyset, adlockpessimistic
    rsact.Open "select r.act_idy,a.act_nme,a.key_act from rdy r,act a where a.grp_clf='Debtor' and a.cst_typ in ('Retailer','Key A/c','Local Key A/c','Regnl Key A/c') and r.act_idy like a.act_idy and a.cat_idy like '" & Trim(txtCategory) & "%' and a.loc_idy like '" & Trim(txtLocation) & "%' and a.mar_idy like '" & Trim(txtMarket) & "%' and r.smn_idy like '" & Trim(txtSalesman) & "%' and r.rou_idy like '" & Trim(cmbRoute) & "%' group by a.act_idy order by a.key_act desc", con, adOpenKeyset, adLockPessimistic
    r = 1
    MS.Rows = 1
    If Not rsact.EOF Then
    MS.Rows = MS.Rows + rsact.RecordCount
    While Not rsact.EOF
        MS.TextMatrix(r, 0) = Trim(rsact(1))
        MS.TextMatrix(r, 14) = Trim(rsact(0))
        MS.TextMatrix(r, 1) = Trim(rsact(2))
        r = r + 1
        rsact.MoveNext
    Wend
    MS.Row = 1
    MS.Col = 2
    MS.SetFocus
    Else
        MsgBox "Retailers not Found!Define Retailers", vbInformation + vbOKOnly, "DAS Version 4.0"
        Exit Sub
    End If
Else
    MOD1 = True
    MS.Rows = 1
    r = 1
    If rsact.State = 1 Then rsact.Close
'    rsact.Open "select act_nme,act_idy from act where cat_idy like '" & Trim(txtCategory) & "%' and loc_idy like '" & Trim(txtLocation) & "%' and mar_idy like '" & Trim(txtMarket) & "%' and smn_idy like '" & Trim(txtSalesman) & "%' and rou_idy like '" & Trim(cmbRoute.Text) & "%'", CON, adOpenKeyset, adlockpessimistic
    rsact.Open "select a.act_nme,r.act_idy,a.key_act from rdy r,act a where a.grp_clf='Debtor' and a.cst_typ in ('Retailer','Key A/c','Local Key A/c','Regnl Key A/c') and a.cat_idy like '" & Trim(txtCategory) & "%' and a.loc_idy like '" & Trim(txtLocation) & "%' and a.mar_idy like '" & Trim(txtMarket) & "%' and r.smn_idy like '" & Trim(txtSalesman) & "%' and r.rou_idy like '" & Trim(cmbRoute) & "%' and a.act_idy like r.act_idy order by a.key_act desc", con, adOpenKeyset, adLockPessimistic
    If Not rsact.EOF Then
        While Not rsRtg.EOF
            If rsact.RecordCount > 0 Then rsact.MoveFirst
            rsact.filter = "act_idy='" & Trim(rsRtg(0)) & "'"
            If Not rsact.EOF Then
                MS.Rows = MS.Rows + 1
                MS.TextMatrix(r, 0) = Trim(rsact(0))
                MS.TextMatrix(r, 14) = Trim(rsact(1))
                MS.TextMatrix(r, 1) = Trim(rsRtg(1))
                MS.TextMatrix(r, 2) = Val(rsRtg(2))
                MS.TextMatrix(r, 3) = Val(rsRtg(3))
                MS.TextMatrix(r, 4) = Val(rsRtg(4))
                MS.TextMatrix(r, 5) = Val(rsRtg(5))
                MS.TextMatrix(r, 6) = Val(rsRtg(6))
                MS.TextMatrix(r, 7) = Val(rsRtg(7))
                MS.TextMatrix(r, 8) = Val(rsRtg(8))
                MS.TextMatrix(r, 9) = Val(rsRtg(9))
                MS.TextMatrix(r, 10) = Val(rsRtg(10))
                MS.TextMatrix(r, 11) = Val(rsRtg(11))
                MS.TextMatrix(r, 12) = Val(rsRtg(12))
                MS.TextMatrix(r, 13) = Val(rsRtg(13))
                r = r + 1
            End If
            rsRtg.MoveNext
        Wend
        MS.SetFocus
    Else
        MsgBox "Retailers not Found! Define Retailer", vbInformation + vbOKOnly, "DAS Version 4.0"
        cancel
        'Command1.SetFocus
    End If
End If
If rsact.State = 1 Then rsact.Close
Set rsact = Nothing
If rsRtg.State = 1 Then rsRtg.Close
Set rsRtg = Nothing
End Sub

Private Sub Form_Activate()
If CHECKTB("rtg.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    Unload Me
    Exit Sub
End If
If txtCategory.Enabled = True Then
    'txtCategory.SetFocus
End If
SSTab1.TabEnabled(1) = False
Check1.SetFocus
End Sub


Private Sub Form_KeyDown(K As Integer, s As Integer)
If K = 120 Then Shell "c:\windows\calc.exe"
If K > 111 And K < 120 Then If TypeOf M_FRM_RTG.ActiveControl Is TextBox Then M_FRM_RTG.ActiveControl.Text = Trim(FUNKEY(K))
If list1.Visible = False Then
    If K = 78 And s = 2 Then
        'newrec
    ElseIf K = 83 And s = 2 Then
        saveTarget
    ElseIf K = 81 And s = 2 Then
        cancel
    ElseIf K = 88 And s = 2 Then
        Unload Me
    ElseIf K = 67 And s = 2 Then
    '   If Toolbar1.Buttons(12).Enabled = True Then CREATE
    ElseIf K = 68 And s = 2 Then
    '    delete
    ElseIf K = 69 And s = 2 Then
        If M_FRM_RTG.ActiveControl.name = "txtCategory" Then channellist
        If M_FRM_RTG.ActiveControl.name = "txtSalesman" Then smnlist
        If M_FRM_RTG.ActiveControl.name = "txtMarket" Then marlist
        If M_FRM_RTG.ActiveControl.name = "txtLocation" Then loclist
    ElseIf K = 112 Then
        M_HELP
    ElseIf K = 80 And s = 2 Then
        'Openreport CURDIR & "\lists.RP1", "Categories list", 0, "", 6
    ElseIf K = 82 And s = 2 Then
        'MDIForm1.RRReport1.ReportName = App.Path & "\Channels list.rrw"
        'MDIForm1.RRReport1.Destination = 1
        'MDIForm1.RRReport1.Printer = "?"
        'MDIForm1.RRReport1.RunReport 1
    End If
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        list1.Visible = False
        SSTab1.Tab = 0
        SSTab1.TabEnabled(1) = False
        If Label2.Caption = "Select Retailer Category" Then
            txtCategory.Enabled = True
            txtCategory.SetFocus
        ElseIf Label2.Caption = "Select Location" Then
            txtLocation.Enabled = True
            txtLocation.SetFocus
        ElseIf Label2.Caption = "Select Salesman" Then
            txtSalesman.Enabled = True
            txtSalesman.SetFocus
        ElseIf Label2.Caption = "Select Market" Then
            txtMarket.Enabled = True
            txtMarket.SetFocus
        End If
    Else
        If M_FRM_RTG.ActiveControl.name = "txtCategory" Then
            Unload Me
        ElseIf M_FRM_RTG.ActiveControl.name = "SSTab1" Then
        Else
            saveTarget
        End If
    End If
End If

End Sub

Private Sub Form_Load()
loadms
rsqry = ""
If rsrdy.State = 1 Then rsrdy.Close
rsrdy.Open "select rou_idy from rdy group by rou_idy", con, adOpenKeyset, adLockPessimistic
cmbRoute.AddItem ""
While Not rsrdy.EOF
    cmbRoute.AddItem Trim(rsrdy(0)) & ""
    rsrdy.MoveNext
Wend
Check1.Value = 1
'Check1.SetFocus
End Sub
Private Sub loadms()
If rsdir.State = 1 Then rsdir.Close
rsdir.Open "select * FROM YER where ACT_PAS = 'Y'", con, adOpenKeyset, adLockPessimistic
If Not rsdir.EOF Then
    rsdir.MoveFirst
    SDATE = Trim(rsdir("FY_PATH"))
    TDATE = Val(Trim(rsdir("FY_PATH"))) + 1
    TDATE = str(TDATE)
End If
With MS
    .Rows = 1
    .Cols = 15
    .TextMatrix(0, 0) = "Party Name"
    .ColWidth(0) = 2000
    .TextMatrix(0, 1) = "Key Act"
    .ColWidth(1) = 1000
    .TextMatrix(0, 2) = "Apr-" & SDATE
    .ColWidth(2) = 1500
    .TextMatrix(0, 3) = "May-" & SDATE
    .ColWidth(3) = 1500
    .TextMatrix(0, 4) = "Jun-" & SDATE
    .ColWidth(4) = 1500
    .TextMatrix(0, 5) = "Jul-" & SDATE
    .ColWidth(5) = 1500
    .TextMatrix(0, 6) = "Aug-" & SDATE
    .ColWidth(6) = 1500
    .TextMatrix(0, 7) = "Sep-" & SDATE
    .ColWidth(7) = 1500
    .TextMatrix(0, 8) = "Oct-" & SDATE
    .ColWidth(8) = 1500
    .TextMatrix(0, 9) = "Nov-" & SDATE
    .ColWidth(9) = 1500
    .TextMatrix(0, 10) = "Dec-" & SDATE
    .ColWidth(10) = 1500
    .TextMatrix(0, 11) = "Jan-" & TDATE
    .ColWidth(11) = 1500
    .TextMatrix(0, 12) = "Feb-" & TDATE
    .ColWidth(12) = 1500
    .TextMatrix(0, 13) = "Mar-" & TDATE
    .ColWidth(13) = 1500
    .ColWidth(14) = 0
    .FixedCols = 1
End With
End Sub

Private Sub MS_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
    If MS.Col <> 0 And MS.Col <> 1 And MS.Col <> 14 And MS.Row > 0 Then
        txtTemp.Visible = True
        txtTemp.Left = MS.Left + MS.CellLeft
        txtTemp.Top = MS.Top + MS.CellTop
        txtTemp.Height = MS.CellHeight
        txtTemp.Width = MS.CellWidth
        txtTemp = Chr(KeyAscii)
        txtTemp.SelStart = Len(Trim(txtTemp))
        txtTemp.SetFocus
    End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "s" Then
    saveTarget
ElseIf Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "c" Then
    cancel
ElseIf Button.KEY = "h" Then
M_HELP
End If
End Sub

Private Sub txtCategory_gotfocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Category ID or Press <Enter> to get List of Category ID's"
End Sub

Private Sub txtCategory_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 27 And txtCategory = "" Then Unload Me
End Sub

Private Sub txtCategory_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If KeyCode = 40 Then
    txtLocation.Enabled = True
    txtLocation.SetFocus
Exit Sub
End If
If KeyCode = 13 Then
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If rs.State = 1 Then rs.Close
    rs.Open "select * from cAT where cAT_idy like '" & Sincrement(txtCategory) & "'", con, adOpenStatic
    If rs.RecordCount > 0 Then
        txtCategory = Sincrement(txtCategory)
        lblCat = Trim(rs("cAT_nme"))
        txtCategory.Enabled = False
        txtLocation.SetFocus
    Else
        channellist
    End If
End If
End Sub
Private Sub channellist()
If CHECKTB("cat.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from CAT where CAT_nme like '" & Trim(txtCategory) & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    Label20 = "Select Retailer Category"
    list1.ListItems.CLEAR
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("CAT_nme")))
        ls.ListSubItems.Add , , Trim(rs("CAT_idy"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.ListItems(1).Selected = True 'Added by PNR
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "Category Not Found, Define Category !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    txtCategory.Enabled = True
    txtCategory.SetFocus
End If
End Sub

Private Sub txtSalesman_gotfocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Salesman ID or Press <Enter> to get List of Salesman ID's"
End Sub

Private Sub txtSalesman_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 27 And txtSalesman = "" Then txtCategory.SetFocus
End Sub

Private Sub txtSalesman_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If KeyCode = 40 Then
    txtMarket.Enabled = True
    txtMarket.SetFocus
Exit Sub
ElseIf KeyCode = 38 And Shift = 1 Then
    txtLocation.Enabled = True
    txtLocation.SetFocus
    Exit Sub
End If
If KeyCode = 13 Then
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If rs.State = 1 Then rs.Close
    rs.Open " select * from smn where smn_idy like '" & Sincrement(txtSalesman) & "'", con, adOpenStatic
    If rs.RecordCount > 0 Then
        txtSalesman = Sincrement(txtSalesman)
        lblSmn = Trim(rs("smn_nme"))
        txtSalesman.Enabled = False
        txtMarket.SetFocus
    Else
        smnlist
    End If
End If
End Sub
Private Sub smnlist()
If CHECKTB("smn.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from smn where smn_nme like '" & Trim(txtSalesman) & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    Label20 = "Select Salesman"
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("smn_nme")))
        ls.ListSubItems.Add , , Trim(rs("smn_idy"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.ListItems(1).Selected = True 'Added by PNR
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    txtSalesman.Enabled = True
    txtSalesman.SetFocus
End If
End Sub


Private Sub txtLocation_gotfocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Location ID or Press <Enter> to get List of Location ID's"
End Sub

Private Sub txtLocation_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 27 And txtLocation = "" Then txtLocation.SetFocus
End Sub

Private Sub txtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If KeyCode = 40 Then
    txtSalesman.Enabled = True
    txtSalesman.SetFocus
Exit Sub
ElseIf KeyCode = 38 And Shift = 1 Then
    txtCategory.Enabled = True
    txtCategory.SetFocus
    Exit Sub
End If
If KeyCode = 13 Or KeyCode = 40 Then
If CHECKTB("loc.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If rs.State = 1 Then rs.Close
    rs.Open " select * from loc where loc_idy like '" & Sincrement(txtLocation) & "'", con, adOpenStatic
    If rs.RecordCount > 0 Then
        txtLocation = Sincrement(txtLocation)
        lblLoc = Trim(rs("loc_nme"))
        txtLocation.Enabled = False
        txtSalesman.SetFocus
    Else
        loclist
    End If
End If
End Sub
Private Sub loclist()
If CHECKTB("loc.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from loc where loc_nme like '" & Trim(txtLocation) & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    Label20 = "Select Location"
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("loc_nme")))
        ls.ListSubItems.Add , , Trim(rs("loc_idy"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.ListItems(1).Selected = True 'Added by PNR
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "Location Not Found, Define Location !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    txtLocation.Enabled = True
    txtLocation.SetFocus
End If
End Sub
Private Sub txtMarket_gotfocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Market ID or Press <Enter> to get List of Market ID's"
End Sub

Private Sub txtMarket_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Or K = 37 Then K = 0
If K = 27 And txtMarket = "" Then txtMarket.SetFocus
End Sub

Private Sub txtMarket_KeyDown(KeyCode As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If KeyCode = 40 Then
    cmbRoute.Enabled = True
    cmbRoute.SetFocus
Exit Sub
ElseIf KeyCode = 38 And Shift = 1 Then
    txtSalesman.Enabled = True
    txtSalesman.SetFocus
    Exit Sub
End If
If KeyCode = 13 Then
If CHECKTB("mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
If rs.State = 1 Then rs.Close
    rs.Open " select * from mar where mar_idy like '" & Sincrement(txtMarket) & "'", con, adOpenStatic
    If rs.RecordCount > 0 Then
        txtMarket = Sincrement(txtMarket)
        lblMar = Trim(rs("mar_nme"))
        txtMarket.Enabled = False
        cmbRoute.SetFocus
    Else
        marlist
    End If
End If
End Sub
Private Sub marlist()
If CHECKTB("mar.dbf") = 1 Then
    MsgBox UCase(TNAME) & " Missing in " & Me.Caption
    cancel
    Exit Sub
End If
Dim ls As ListItem
Dim rs As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "select * from mar where mar_nme like '" & Trim(txtMarket) & "%'", con, adOpenKeyset, adLockPessimistic
If rs.RecordCount > 0 Then
    list1.ListItems.CLEAR
    Label20 = "Select Market"
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("mar_nme")))
        ls.ListSubItems.Add , , Trim(rs("mar_idy"))
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.ListItems(1).Selected = True 'Added by PNR
    list1.Visible = True
    list1.SetFocus
Else
    MsgBox "Market Not Found, Define Market !", vbExclamation, "DAS Version 4.0"
    SSTab1.Tab = 0
    txtMarket.Enabled = True
    txtMarket.SetFocus
End If
End Sub

Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Category or Press <Esc> to Quit"
End Sub

Private Sub list1_KeyDown(K As Integer, s As Integer)
Dim rs As New ADODB.Recordset
Dim lst As ListSubItem
If K = 13 Then
    If list1.ListItems.count > 0 Then
        If Label20 = "Select Retailer Category" Then
            Set lst = list1.SelectedItem.ListSubItems.Item(1)
            If rs.State = 1 Then rs.Close
            rs.Open "select * from CAT where CAT_idy like '" & Trim(lst.Text) & "'", con, adOpenStatic
            txtCategory = Trim(rs("CAT_idy"))
            lblCat = Trim(rs("CAT_nme"))
            list1.Visible = False
            SSTab1.TabEnabled(1) = False
            SSTab1.Tab = 0
            txtCategory.Enabled = False
            txtLocation.Enabled = True
            txtLocation.SetFocus
        ElseIf Label20 = "Select Location" Then
            Set lst = list1.SelectedItem.ListSubItems.Item(1)
            If rs.State = 1 Then rs.Close
            rs.Open "select * from loc where loc_idy like '" & Trim(lst.Text) & "'", con, adOpenStatic
            txtLocation = Trim(rs("loc_idy"))
            lblLoc = Trim(rs("loc_nme"))
            list1.Visible = False
            SSTab1.TabEnabled(1) = False
            SSTab1.Tab = 0
            txtLocation.Enabled = False
            txtSalesman.SetFocus
        ElseIf Label20 = "Select Salesman" Then
            Set lst = list1.SelectedItem.ListSubItems.Item(1)
            If rs.State = 1 Then rs.Close
            rs.Open "select * from smn where smn_idy like '" & Trim(lst.Text) & "'", con, adOpenStatic
            txtSalesman = Trim(rs("smn_idy"))
            lblSmn = Trim(rs("smn_nme"))
            list1.Visible = False
            SSTab1.TabEnabled(1) = False
            SSTab1.Tab = 0
            txtSalesman.Enabled = False
            txtMarket.SetFocus
        ElseIf Label20 = "Select Market" Then
            Set lst = list1.SelectedItem.ListSubItems.Item(1)
            If rs.State = 1 Then rs.Close
            rs.Open "select * from mar where mar_idy like '" & Trim(lst.Text) & "'", con, adOpenStatic
            txtMarket = Trim(rs("mar_idy"))
            lblMar = Trim(rs("mar_nme"))
            list1.Visible = False
            SSTab1.TabEnabled(1) = False
            SSTab1.Tab = 0
            txtMarket.Enabled = False
            cmbRoute.SetFocus
        End If
    End If
End If
End Sub
Private Sub cancel()
For Each Control In Me.Controls
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
lblCat = ""
lblSmn = ""
lblMar = ""
lblLoc = ""
txtCategory.Enabled = True
txtSalesman.Enabled = True
txtLocation.Enabled = True
txtMarket.Enabled = True
cmbRoute.ListIndex = 0
If txtTemp.Visible = True Then txtTemp.Visible = False
txtCategory.Enabled = True
txtCategory.SetFocus
MS.Rows = 1
Check1.Value = 0
End Sub
Private Sub saveTarget()
If MsgBox("Confirm Targets Save?", vbQuestion + vbYesNo, "DAS Version 4.0") = vbNo Then
    cancel
    Exit Sub
End If
If MS.Rows > 1 Then
    If Len(Trim(MS.TextMatrix(1, 0))) > 0 Then
    Else
    MsgBox "Targets not Found!", vbInformation + vbOKOnly, "DAS Version 4.0"
    Command1.Enabled = True
    Command1.SetFocus
    Exit Sub
    End If
Else
    MsgBox "Targets not Found!", vbInformation + vbOKOnly, "DAS Version 4.0"
    Command1.Enabled = True
    Command1.SetFocus
    Exit Sub
End If

If rsRtg.State = 1 Then rsRtg.Close
rsRtg.Open "select * from rtg", con, adOpenKeyset, adLockPessimistic
If MOD1 = True Then
    For r = 1 To MS.Rows - 1
        If rsRtg.RecordCount > 0 Then rsRtg.MoveFirst
        rsRtg.filter = "act_idy='" & Trim(MS.TextMatrix(r, 14)) & "'"
        If Not rsRtg.EOF Then
            rsRtg(0) = Trim(MS.TextMatrix(r, 14))
            rsRtg(1) = Trim(MS.TextMatrix(r, 1))
            rsRtg(2) = Val(Trim(MS.TextMatrix(r, 2)))
            rsRtg(3) = Val(Trim(MS.TextMatrix(r, 3)))
            rsRtg(4) = Val(Trim(MS.TextMatrix(r, 4)))
            rsRtg(5) = Val(Trim(MS.TextMatrix(r, 5)))
            rsRtg(6) = Val(Trim(MS.TextMatrix(r, 6)))
            rsRtg(7) = Val(Trim(MS.TextMatrix(r, 7)))
            rsRtg(8) = Val(Trim(MS.TextMatrix(r, 8)))
            rsRtg(9) = Val(Trim(MS.TextMatrix(r, 9)))
            rsRtg(10) = Val(Trim(MS.TextMatrix(r, 10)))
            rsRtg(11) = Val(Trim(MS.TextMatrix(r, 11)))
            rsRtg(12) = Val(Trim(MS.TextMatrix(r, 12)))
            rsRtg(13) = Val(Trim(MS.TextMatrix(r, 13)))
            rsRtg(14) = Val(SDATE)
            rsRtg.Update
        End If
    Next
Else
   For r = 1 To MS.Rows - 1
       rsRtg.AddNew
       rsRtg(0) = Trim(MS.TextMatrix(r, 14))
       rsRtg(1) = Trim(MS.TextMatrix(r, 1))
       rsRtg(2) = Val(Trim(MS.TextMatrix(r, 2)))
       rsRtg(3) = Val(Trim(MS.TextMatrix(r, 3)))
       rsRtg(4) = Val(Trim(MS.TextMatrix(r, 4)))
       rsRtg(5) = Val(Trim(MS.TextMatrix(r, 5)))
       rsRtg(6) = Val(Trim(MS.TextMatrix(r, 6)))
       rsRtg(7) = Val(Trim(MS.TextMatrix(r, 7)))
       rsRtg(8) = Val(Trim(MS.TextMatrix(r, 8)))
       rsRtg(9) = Val(Trim(MS.TextMatrix(r, 9)))
       rsRtg(10) = Val(Trim(MS.TextMatrix(r, 10)))
       rsRtg(11) = Val(Trim(MS.TextMatrix(r, 11)))
       rsRtg(12) = Val(Trim(MS.TextMatrix(r, 12)))
       rsRtg(13) = Val(Trim(MS.TextMatrix(r, 13)))
       rsRtg(14) = Val(SDATE)
       rsRtg.Update
   Next
End If
If rsRtg.State = 1 Then rsRtg.Close
Set rsRtg = Nothing
cancel
End Sub

Private Sub txttemp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    MS.Text = Val(txtTemp)
    txtTemp.Visible = False
    If MS.Col <> 13 Then
        MS.Col = MS.Col + 1
    Else
        If MS.Row <> MS.Rows - 1 Then
            MS.Row = MS.Row + 1
            MS.Col = 2
        End If
    End If
    MS.SetFocus
End If
End Sub

Private Sub txtTemp_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Or KeyAscii = 42 Then
Else
    KeyAscii = 0
End If
End Sub
