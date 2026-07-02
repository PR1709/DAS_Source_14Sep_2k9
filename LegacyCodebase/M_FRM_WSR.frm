VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form M_FRM_WSR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Weekly Sales Report"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "M_FRM_WSR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5480
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   400
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9657
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Weekly Sales Report"
      TabPicture(0)   =   "M_FRM_WSR.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Combo1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "MS"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_WSR.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "list1"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5685
         TabIndex        =   20
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3360
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Export"
         Height          =   375
         Left            =   6120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Generate"
         Height          =   375
         Left            =   7200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1680
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MS 
         Height          =   3255
         Left            =   60
         TabIndex        =   12
         Top             =   2160
         Width           =   8170
         _ExtentX        =   14420
         _ExtentY        =   5741
         _Version        =   393216
         Rows            =   3
         Cols            =   16
         FixedRows       =   2
         RowHeightMin    =   330
         BackColor       =   -2147483633
         BackColorSel    =   16744576
         FocusRect       =   2
         GridLines       =   2
         MergeCells      =   2
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   4200
         TabIndex        =   8
         Top             =   360
         Width           =   3015
         Begin VB.ComboBox ROU_LST 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1200
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   652
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&Salesman"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&Route"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   3120
            TabIndex        =   17
            Top             =   720
            Visible         =   0   'False
            Width           =   1200
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sales Calender"
         Height          =   1095
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   2055
         Begin VB.Label Enddate 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   600
            TabIndex        =   7
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Start 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   600
            TabIndex        =   6
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "End"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   675
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Start"
            Height          =   195
            Left            =   75
            TabIndex        =   4
            Top             =   308
            Width           =   330
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
      End
      Begin MSComctlLib.ListView list1 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   7646
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&File Name"
         Height          =   195
         Left            =   2520
         TabIndex        =   19
         Top             =   1755
         Width           =   705
      End
      Begin VB.Label Label10 
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
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   -74880
         TabIndex        =   16
         Top             =   720
         Width           =   8055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Week No"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   660
         Width           =   690
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
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
            Picture         =   "M_FRM_WSR.frx":047A
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":09BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":0D86
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":10DA
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":1712
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":1A66
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":1EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":21BA
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":25C6
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":26D2
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":2A26
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":2E3A
            Key             =   "cpy"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_WSR.frx":3386
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   8280
      _ExtentX        =   14605
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
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "WSR"
                  Text            =   "Weekly Sales Report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EWSR"
                  Text            =   "Extended Week Report"
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
End
Attribute VB_Name = "M_FRM_WSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WEK As New ADODB.Recordset
Dim FPATH As String

Private Sub Combo1_Change()
WEK.filter = "WEK_IDY=" & Val(Combo1)
If WEK.RecordCount > 0 Then
Start = Format(WEK!FRM_DAT, "DD/MMM/YYYY")
ENDDATE = Format(WEK!TO_DAT, "DD/MMM/YYYY")
End If
End Sub

Private Sub Combo1_Click()
    WEK.filter = "WEK_IDY=" & Val(Combo1)
    If WEK.RecordCount > 0 Then
    Start = Format(WEK!FRM_DAT, "DD/MMM/YYYY")
    ENDDATE = Format(WEK!TO_DAT, "DD/MMM/YYYY")
    End If
End Sub

Private Sub Combo1_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Combo1) > 0 And Val(Combo1) < 53 Then
        Option1.SetFocus
    Else
        MsgBox "Invalid FMCG Calendar WEEK Selection", vbExclamation, "DAS Version 4.0"
        Combo1.SetFocus
    End If
End If
End Sub

Private Sub Command1_Click()
If Val(Combo1) > 0 And Val(Combo1) < 53 Then
Else
    MsgBox "Invalid FMCG Calendar WEEK Selection", vbExclamation, "DAS Version 4.0"
    Combo1.SetFocus
    Exit Sub
End If
If IsDate(Start) = False Then
    MsgBox "Define FMCG Calendar Dates to generate report!", vbExclamation, "DAS Version 4.0"
    Combo1.SetFocus
    Exit Sub
End If
Me.MousePointer = vbHourglass
MS.Rows = 2
MS.Cols = 16
MS.Refresh
For A = 1 To MS.Cols - 1
    If A Mod 2 <> 0 Then
        MS.TextMatrix(1, A) = "Sales"
    Else
        MS.TextMatrix(1, A) = "Cls.Stk"
    End If
Next
MS.MergeRow(0) = True
MS.TextMatrix(0, 1) = Start
MS.TextMatrix(0, 2) = Start
If DateDiff("D", DateAdd("D", 1, Start), ENDDATE) >= 0 Then
    MS.TextMatrix(0, 3) = Format(DateAdd("D", 1, Start), "DD/MMM/YYYY")
    MS.TextMatrix(0, 4) = Format(DateAdd("D", 1, Start), "DD/MMM/YYYY")
End If
If DateDiff("D", DateAdd("D", 2, Start), ENDDATE) >= 0 Then
    MS.TextMatrix(0, 5) = Format(DateAdd("D", 2, Start), "DD/MMM/YYYY")
    MS.TextMatrix(0, 6) = Format(DateAdd("D", 2, Start), "DD/MMM/YYYY")
End If
If DateDiff("D", DateAdd("D", 3, Start), ENDDATE) >= 0 Then
    MS.TextMatrix(0, 7) = Format(DateAdd("D", 3, Start), "DD/MMM/YYYY")
    MS.TextMatrix(0, 8) = Format(DateAdd("D", 3, Start), "DD/MMM/YYYY")
End If
If DateDiff("D", DateAdd("D", 4, Start), ENDDATE) >= 0 Then
    MS.TextMatrix(0, 9) = Format(DateAdd("D", 4, Start), "DD/MMM/YYYY")
    MS.TextMatrix(0, 10) = Format(DateAdd("D", 4, Start), "DD/MMM/YYYY")
End If
If DateDiff("D", DateAdd("D", 5, Start), ENDDATE) >= 0 Then
    MS.TextMatrix(0, 11) = Format(DateAdd("D", 5, Start), "DD/MMM/YYYY")
    MS.TextMatrix(0, 12) = Format(DateAdd("D", 5, Start), "DD/MMM/YYYY")
End If
If DateDiff("D", DateAdd("D", 6, Start), ENDDATE) >= 0 Then
    MS.TextMatrix(0, 13) = Format(DateAdd("D", 6, Start), "DD/MMM/YYYY")
    MS.TextMatrix(0, 14) = Format(DateAdd("D", 6, Start), "DD/MMM/YYYY")
End If
MS.TextMatrix(1, 15) = "Product ID"
Dim rs As New ADODB.Recordset
MS.Rows = 2
For A = 1 To MS.Cols - 2 Step 2
    If IsDate(MS.TextMatrix(0, A)) = True Then
        WSRUPDATECLS MS.TextMatrix(0, A)
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM CLD,PRD,BRD WHERE CLD.PRD_IDY LIKE BRD.BAT_IDY AND BRD.PRD_IDY LIKE PRD.PRD_IDY GROUP BY PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
        For b = 1 To rs.RecordCount
            For C = 1 To MS.Rows - 1
                If Trim(MS.TextMatrix(C, 15)) = Trim(rs!prd_idy) Then
                    MS.TextMatrix(C, A + 1) = QTY(STM(Val(rs!cls_stk & ""), Val(rs!SUB_UNT)))
                    GoTo LBL:
                End If
            Next
            MS.AddItem ""
            MS.TextMatrix(MS.Rows - 1, 0) = Trim(rs!PRD_NME)
            MS.TextMatrix(MS.Rows - 1, A + 1) = QTY(STM(Val(rs!cls_stk & ""), Val(rs!SUB_UNT)))
            MS.TextMatrix(MS.Rows - 1, 15) = Trim(rs!prd_idy)
LBL:
            rs.MoveNext
        Next
    End If
Next

For A = 1 To MS.Cols - 2 Step 2
    If IsDate(MS.TextMatrix(0, A)) = True Then
        WSRUPDATESAL MS.TextMatrix(0, A)
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM CLD,PRD,BRD WHERE CLD.PRD_IDY LIKE BRD.BAT_IDY AND BRD.PRD_IDY LIKE PRD.PRD_IDY GROUP BY PRD.PRD_IDY", con, adOpenKeyset, adLockPessimistic
        For b = 1 To rs.RecordCount
            For C = 1 To MS.Rows - 1
                If Trim(MS.TextMatrix(C, 15)) = Trim(rs!prd_idy) Then
                    MS.TextMatrix(C, A) = QTY(STM(Val(rs!sal_stk & ""), Val(rs!SUB_UNT)))
                    GoTo LBL1:
                End If
            Next
            MS.AddItem ""
            MS.TextMatrix(MS.Rows - 1, 0) = Trim(rs!PRD_NME)
            MS.TextMatrix(MS.Rows - 1, A) = QTY(STM(Val(rs!sal_stk & ""), Val(rs!SUB_UNT)))
            MS.TextMatrix(MS.Rows - 1, 15) = Trim(rs!prd_idy)
LBL1:
            rs.MoveNext
        Next
    End If
Next

A1 = 0
A2 = 0
A3 = 0
A4 = 0
A5 = 0
A6 = 0
A7 = 0

For A = 2 To MS.Rows - 1
    A1 = Val(A1) + Val(MS.TextMatrix(A, 1))
    A2 = Val(A2) + Val(MS.TextMatrix(A, 3))
    A3 = Val(A3) + Val(MS.TextMatrix(A, 5))
    A4 = Val(A4) + Val(MS.TextMatrix(A, 7))
    A5 = Val(A5) + Val(MS.TextMatrix(A, 9))
    A6 = Val(A6) + Val(MS.TextMatrix(A, 11))
    A7 = Val(A7) + Val(MS.TextMatrix(A, 13))
Next
MS.AddItem ""
MS.TextMatrix(MS.Rows - 1, 0) = "Total"
MS.TextMatrix(MS.Rows - 1, 1) = QTY((Val(A1)))
MS.TextMatrix(MS.Rows - 1, 3) = QTY(Val(A2))
MS.TextMatrix(MS.Rows - 1, 5) = QTY(Val(A3))
MS.TextMatrix(MS.Rows - 1, 7) = QTY(Val(A4))
MS.TextMatrix(MS.Rows - 1, 9) = QTY(Val(A5))
MS.TextMatrix(MS.Rows - 1, 11) = QTY(Val(A6))
MS.TextMatrix(MS.Rows - 1, 13) = QTY(Val(A7))
MS.Row = MS.Rows - 1
For A = 0 To MS.Cols - 1
    MS.Col = A
    MS.CellFontBold = True
Next
MS.Cols = MS.Cols + 1
MS.TextMatrix(1, MS.Cols - 1) = "Sales Total"
MS.Col = MS.Cols - 1
For A = 2 To MS.Rows - 1
    MS.TextMatrix(A, MS.Cols - 1) = QTY(Val(MS.TextMatrix(A, 1)) + Val(MS.TextMatrix(A, 3)) + Val(MS.TextMatrix(A, 5)) + Val(MS.TextMatrix(A, 7)) + Val(MS.TextMatrix(A, 9)) + Val(MS.TextMatrix(A, 11)) + Val(MS.TextMatrix(A, 13)))
    MS.Row = A
    MS.CellFontBold = True
Next
Me.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
If Trim(FPATH) = "" And Text2 = "" Then
    MsgBox "Invalid File Name", vbExclamation, "DAS Version 4.0"
    Text2.SetFocus
    Exit Sub
End If
Open Text2 For Output As #1   'Open file for output.
STR1 = ""
For r = 0 To MS.Rows - 1
    For C = 0 To MS.Cols - 1
        If r = 0 And C = 0 Then
            STR1 = STR1 & ","
        Else
            STR1 = STR1 & Trim(MS.TextMatrix(r, C)) & ","
        End If
    Next
    Write #1, STR1
    STR1 = ""
Next
Close #1
MsgBox "Data Exported to " & Text2, vbExclamation, "DAS Version 4.0"
End Sub

Private Sub Command3_Click()
CommonDialog1.filter = "*.TXT"
CommonDialog1.DefaultExt = "TXT"
CommonDialog1.DialogTitle = "Enter File Name"
CommonDialog1.ShowSave
Text2 = CommonDialog1.FileName
End Sub

Private Sub Command4_Click()
Dim rs As New ADODB.Recordset
If Dir(CURDIR & "\wsr.dbf") = "" Then
    con.Execute "Create table wsr(prd_NME char(30),day_1 number(15,3) NULL,day_2 number(15,3) NULL,day_3 number(15,3) NULL,day_4  number(15,3) NULL,day_5 number(15,3) NULL,day_6 number(15,3) NULL,day_7 number(15,3) NULL,day_8 number(15,3) NULL,day_9 number(15,3)NULL,day_10 number(15,3) NULL,day_11 number(15,3) NULL)"
End If
con.Execute "DELETE FROM WSR"
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM WSR", con, adOpenKeyset, adLockPessimistic
For A = 2 To MS.Rows - 1
    rs.AddNew
    rs!PRD_NME = MS.TextMatrix(A, 0)
    rs!DAY_1 = Val(MS.TextMatrix(A, 2))
    rs!DAY_2 = Val(MS.TextMatrix(A, 4))
    rs!DAY_3 = Val(MS.TextMatrix(A, 6))
    rs!DAY_4 = Val(MS.TextMatrix(A, 8))
    rs!DAY_5 = Val(MS.TextMatrix(A, 10))
    rs!DAY_6 = Val(MS.TextMatrix(A, 12))
    rs!DAY_7 = Val(MS.TextMatrix(A, 14))
    rs.Update
    rs.AddNew
    rs!PRD_NME = ""
    rs!DAY_1 = Val(MS.TextMatrix(A, 2 - 1))
    rs!DAY_2 = Val(MS.TextMatrix(A, 4 - 1))
    rs!DAY_3 = Val(MS.TextMatrix(A, 6 - 1))
    rs!DAY_4 = Val(MS.TextMatrix(A, 8 - 1))
    rs!DAY_5 = Val(MS.TextMatrix(A, 10 - 1))
    rs!DAY_6 = Val(MS.TextMatrix(A, 12 - 1))
    rs!DAY_7 = Val(MS.TextMatrix(A, 14 - 1))
    rs.Update
Next
End Sub

Private Sub Form_Activate()
SSTab1.TabEnabled(1) = False
SSTab1.Tab = 0
MS.ColWidth(15) = 0
MS.ColWidth(0) = 2000
MS.MergeCol(0) = True
MS.TextMatrix(0, 0) = "Product Name"
MS.TextMatrix(1, 0) = "Product Name"
Combo1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 88 And Shift = 2 Then Unload Me
If KeyCode = 81 And Shift = 2 Then CANCELRECORD
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If list1.Visible = True Then
        SSTab1.Tab = 0
        list1.Visible = False
        Text1.SetFocus
    Else
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
'ROU_LST.Selected(0) = True
Combo1.CLEAR
For A = 1 To 52
Combo1.AddItem A
Next
AddData ROU_LST, Me, "ROU_IDY", "RDY", "", "GROUP BY ROU_IDY"
ROU_LST.AddItem "All"
ROU_LST.Text = "All"
If WEK.State = 1 Then WEK.Close
WEK.Open "SELECT * FROM WEK", con, adOpenKeyset
End Sub

Private Sub List1_ColumnClick(ByVal CH As MSComctlLib.ColumnHeader)
If list1.SortOrder = lvwAscending Then
list1.SortOrder = lvwDescending
ElseIf list1.SortOrder = lvwDescending Then
list1.SortOrder = lvwAscending
End If
list1.SortKey = CH.Index - 1
End Sub
Private Sub list1_KeyDown(KeyAscii As Integer, s As Integer)
Dim rs As New ADODB.Recordset
If KeyAscii = 13 Then
    If list1.ListItems.count > 0 Then
        If rs.State = 1 Then rs.Close
        Text1 = list1.SelectedItem
        Label7 = list1.SelectedItem.ListSubItems.Item(1)
        list1.Visible = False
        Command1.SetFocus
        SSTab1.Tab = 0
    End If
End If
End Sub

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select ID or Press <Esc> to Quit"
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then ROU_LST.SetFocus
End Sub

Private Sub ROU_LST_kEYPRESS(K As Integer)
If K = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Salesmen ID or Press <Enter> To Get The List of Salesmen ID's"
End Sub

Private Sub Text1_KeyPress(K As Integer)
If K > 96 And K < 123 Then K = K - 32
If K = 39 Or K = 41 Then K = 0
If K = 27 And Text1 = "" Then Unload Me
End Sub

Private Sub TEXT1_KeyDown(K As Integer, Shift As Integer)
Dim rs As New ADODB.Recordset
If K = 13 Or K = 40 Then
    If rs.State = 1 Then rs.Close
    rs.Open "select * from smn where smn_idy like '" & Sincrement(Text1) & "'", con, adOpenStatic
    If rs.RecordCount > 0 Then
        Label7 = Sincrement(Text1)
        Text1 = Trim(rs("smn_nme"))
        Command1.SetFocus
    Else
        smnlist
    End If
End If
End Sub

Private Sub smnlist()
Dim ls As ListItem
Dim rs As New ADODB.Recordset
rs.Open "select * from smn where smn_nme like '" & Text1 & "%'", con, adOpenStatic
list1.ListItems.CLEAR
If rs.RecordCount > 0 Then
    For A = 1 To rs.RecordCount
        Set ls = list1.ListItems.Add(, , Trim(rs("smn_nme")))
        ls.ListSubItems.Add , , Trim(rs("smn_idy") & "")
        ls.ListSubItems.Add , , Trim(rs("smn_ad4") & "")
        rs.MoveNext
    Next
    SSTab1.Tab = 1
    list1.Visible = True
    list1.SetFocus
Else
    If SSTab1.Tab = 0 Then
        MsgBox "Salesman Not Found, Define Salesman !", vbExclamation, "DAS Version 4.0"
        SSTab1.Tab = 0
        Text1.Enabled = True
        Text1.SetFocus
    End If
End If
End Sub


Public Sub WSRUPDATECLS(DAT As Date)
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
con.Execute "DELETE FROM CLD"
rs.Open "select * from cld", con, adOpenKeyset, adLockPessimistic
If RS1.State = 1 Then RS1.Close
'RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat <= ctod('" & DAT & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adlockpessimistic
'RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat <= ctod('" & DAT & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adlockpessimistic 'Modified by Radhika & Ramanesh
'For A = 1 To RS1.RecordCount
'    If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
'        CON.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
'        If d = 0 Then
'            RS.AddNew
'            RS!PRD_IDY = Trim(RS1!BAT_IDY)
'            RS!opn_stk = 0
'            RS!PUR_STK = 0
'            RS!srn_stk = 0
'            RS!sal_stk = Val(RS1(0))
'            RS!PRT_STK = 0
'            RS!STI_STK = 0
'            RS(7) = 0
'            RS!cls_stk = 0
'            RS.Update
'        End If
'    ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
'        CON.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
'        If d = 0 Then
'            RS.AddNew
'            RS!PRD_IDY = Trim(RS1!BAT_IDY)
'            RS!opn_stk = 0
'            RS!PUR_STK = 0
'            RS!srn_stk = 0
'            RS!sal_stk = Val(RS1(0)) * -1
'            RS!PRT_STK = 0
'            RS!STI_STK = 0
'            RS(7) = 0
'            RS!cls_stk = 0
'            RS.Update
'        End If
'    ElseIf RS1!TXN_TYP = "OPS" Then
'        CON.Execute "UPDATE CLD SET OPN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
'        If d = 0 Then
'            RS.AddNew
'            RS!PRD_IDY = Trim(RS1!BAT_IDY)
'            RS!sal_stk = 0
'            RS!PUR_STK = 0
'            RS!srn_stk = 0
'            RS!opn_stk = Val(RS1(0))
'            RS!PRT_STK = 0
'            RS!STI_STK = 0
'            RS(7) = 0
'            RS!cls_stk = 0
'            RS.Update
'        End If
'    ElseIf RS1!TXN_TYP = "PUR" Then
'    CON.Execute "UPDATE CLD SET PUR_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
'        If d = 0 Then
'            RS.AddNew
'            RS!PRD_IDY = Trim(RS1!BAT_IDY)
'            RS!sal_stk = 0
'            RS!opn_stk = 0
'            RS!srn_stk = 0
'            RS!PUR_STK = Val(RS1(0))
'            RS!PRT_STK = 0
'            RS!STI_STK = 0
'            RS(7) = 0
'            RS!cls_stk = 0
'            RS.Update
'        End If
'    ElseIf RS1!TXN_TYP = "PRT" Then
'    CON.Execute "UPDATE CLD SET PRT_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
'        If d = 0 Then
'            RS.AddNew
'            RS!PRD_IDY = Trim(RS1!BAT_IDY)
'            RS!sal_stk = 0
'            RS!opn_stk = 0
'            RS!srn_stk = 0
'            RS!PRT_STK = Val(RS1(0))
'            RS!PUR_STK = 0
'            RS!STI_STK = 0
'            RS(7) = 0
'            RS!cls_stk = 0
'            RS.Update
'        End If
'    ElseIf RS1!TXN_TYP = "STI" Then
'    CON.Execute "UPDATE CLD SET STI_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
'        If d = 0 Then
'            RS.AddNew
'            RS!PRD_IDY = Trim(RS1!BAT_IDY)
'            RS!sal_stk = 0
'            RS!opn_stk = 0
'            RS!srn_stk = 0
'            RS!STI_STK = Val(RS1(0))
'            RS!PRT_STK = 0
'            RS!PUR_STK = 0
'            RS(7) = 0
'            RS!cls_stk = 0
'            RS.Update
'        End If
'    ElseIf RS1!TXN_TYP = "SRN" Then
'    CON.Execute "UPDATE CLD SET SRN_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
'        If d = 0 Then
'            RS.AddNew
'            RS!PRD_IDY = Trim(RS1!BAT_IDY)
'            RS!sal_stk = 0
'            RS!opn_stk = 0
'            RS!PUR_STK = 0
'            RS!srn_stk = Val(RS1(0))
'            RS!PRT_STK = 0
'            RS!STI_STK = 0
'            RS(7) = 0
'            RS!cls_stk = 0
'            RS.Update
'        End If
'    ElseIf RS1!TXN_TYP = "STO" Then
'    CON.Execute "UPDATE CLD SET STO_STK = " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
'        If d = 0 Then
'            RS.AddNew
'            RS!PRD_IDY = Trim(RS1!BAT_IDY)
'            RS!sal_stk = 0
'            RS!opn_stk = 0
'            RS!srn_stk = 0
'            RS(7) = Val(RS1(0))
'            RS!PRT_STK = 0
'            RS!STI_STK = 0
'            RS!PUR_STK = 0
'            RS!cls_stk = 0
'            RS.Update
'        End If
'    End If
'    RS1.MoveNext
'Next
'RS.Close
If RS1.State = 1 Then RS1.Close
'RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where (STK_UPD ='Y' or ORD_IDY # ' ') AND inv_dat < ctod('" & DTPicker1 & "') group by bat_idy,txn_typ,stk_typ", CON, adOpenKeyset, adlockpessimistic
RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat <= ctod('" & DAT & "') group by bat_idy,txn_typ,stk_typ", con, adOpenKeyset, adLockPessimistic
For A = 1 To RS1.RecordCount
    If RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
        con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!sal_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "SAL" And (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
        con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!sal_stk = Val(RS1(0)) * -1
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "OPS" Then
        con.Execute "UPDATE CLD SET OPN_STK = OPN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!opn_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "PUR" Then
    con.Execute "UPDATE CLD SET PUR_STK = PUR_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!PUR_STK = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "PRT" Then
    con.Execute "UPDATE CLD SET PRT_STK = PRT_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!PRT_STK = Val(RS1(0))
            rs!PUR_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "STI" Then
    con.Execute "UPDATE CLD SET STI_STK = STI_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs!STI_STK = Val(RS1(0))
            rs!PRT_STK = 0
            rs!PUR_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "SRN" Then
    con.Execute "UPDATE CLD SET SRN_STK = SRN_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf RS1!TXN_TYP = "STO" Or RS1!TXN_TYP = "ADJ" Then
    con.Execute "UPDATE CLD SET STO_STK = STO_STK + " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!sal_stk = 0
            rs!opn_stk = 0
            rs!srn_stk = 0
            rs(7) = Val(RS1(0))
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs!PUR_STK = 0
            rs!cls_stk = 0
            rs.Update
        End If
    End If
    RS1.MoveNext
Next
con.Execute "UPDATE CLD SET CLS_STK=(PUR_STK+OPN_STK+SRN_STK+STI_STK)-(SAL_STK+PRT_STK+STO_STK)"
End Sub

Public Sub WSRUPDATESAL(DAT As Date)
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
con.Execute "DELETE FROM CLD"
rs.Open "select * from cld", con, adOpenKeyset, adLockPessimistic
If RS1.State = 1 Then RS1.Close
If Option1.Value = True Then
    If ROU_LST = "All" Then
        'RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat = ctod('" & DAT & "') AND TXN_TYP LIKE 'SAL' group by bat_idy,txn_typ", CON, adOpenKeyset, adlockpessimistic
        RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat = ctod('" & DAT & "') AND TXN_TYP LIKE 'SAL' group by bat_idy,txn_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh & Radhika
    Else
        'RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat = ctod('" & DAT & "') AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ROU_IDY LIKE '" & Trim(ROU_LST) & "' AND TXN_TYP LIKE 'SAL') group by bat_idy,txn_typ", CON, adOpenKeyset, adlockpessimistic
        RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat = ctod('" & DAT & "') AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE ROU_IDY LIKE '" & Trim(ROU_LST) & "' AND TXN_TYP LIKE 'SAL') group by bat_idy,txn_typ", con, adOpenKeyset, adLockPessimistic 'Ramanesh & Radhika
    End If
ElseIf Option2.Value = True Then
    'RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat = ctod('" & DAT & "') AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE SMN_IDY LIKE '" & Trim(Label7) & "' AND TXN_TYP LIKE 'SAL') group by bat_idy,txn_typ", CON, adOpenKeyset, adlockpessimistic
    RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat = ctod('" & DAT & "') AND INV_IDY IN(SELECT INV_IDY FROM INL WHERE SMN_IDY LIKE '" & Trim(Label7) & "' AND TXN_TYP LIKE 'SAL') group by bat_idy,txn_typ", con, adOpenKeyset, adLockPessimistic 'Modified by Ramanesh & Radhika
Else
    'RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where STK_UPD ='Y' AND inv_dat = ctod('" & DAT & "') AND TXN_TYP LIKE 'SAL' group by bat_idy,txn_typ", CON, adOpenKeyset, adlockpessimistic
    RS1.Open "select sum(tot_pcs),BAT_IDY,txn_typ,stk_typ from inv where inv_dat = ctod('" & DAT & "') AND TXN_TYP LIKE 'SAL' group by bat_idy,txn_typ", con, adOpenKeyset, adLockPessimistic
End If
For A = 1 To RS1.RecordCount
    If (Trim(RS1!stk_typ) = "Salable" Or Trim(RS1!stk_typ) = "Free") Then
        con.Execute "UPDATE CLD SET SAL_STK = SAL_STK + " & Val(RS1(0) & "") & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!sal_stk = Val(RS1(0) & "")
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    ElseIf (Trim(RS1!stk_typ) = "TakeBack" Or Trim(RS1!stk_typ) = "Damage" Or Trim(RS1!stk_typ) = "Breakage" Or Trim(RS1!stk_typ) = "Leakage") Then
        con.Execute "UPDATE CLD SET SAL_STK = SAL_STK - " & Val(RS1(0)) & " WHERE PRD_IDY LIKE '" & Trim(RS1!BAT_IDY) & "'", d
        If d = 0 Then
            rs.AddNew
            rs!prd_idy = Trim(RS1!BAT_IDY)
            rs!opn_stk = 0
            rs!PUR_STK = 0
            rs!srn_stk = 0
            rs!sal_stk = Val(RS1(0) & "") * -1
            rs!PRT_STK = 0
            rs!STI_STK = 0
            rs(7) = 0
            rs!cls_stk = 0
            rs.Update
        End If
    End If
    RS1.MoveNext
Next
rs.Close
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Or KeyAscii = 41 Then KeyAscii = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.KEY = "e" Then
    Unload Me
ElseIf Button.KEY = "c" Then
    CANCELRECORD
ElseIf Button.KEY = "h" Then
    M_HELP
ElseIf Button.KEY = "p" Then

End If
End Sub

Public Sub CANCELRECORD()
MS.Rows = 2
MS.CLEAR
MS.TextMatrix(0, 0) = "Product Name"
MS.TextMatrix(1, 0) = "Product Name"
Text1 = ""
Text2 = ""
Start = ""
ENDDATE = ""
ROU_LST.Text = "All"
Combo1 = ""
Combo1.SetFocus
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Command4_Click
con.Execute "update dir set DIR_ALS= '" & Format(Start, "MM/DD/YYYY") & "' WHERE dir_idy LIKE 'HDR' AND DIR_IDX='DAT'"
con.Execute "update dir set DIR_MSG= '" & Val(Combo1) & "' WHERE dir_idy LIKE 'HDR' AND DIR_IDX='DAT'"
If ButtonMenu.KEY = "WSR" Then
    OpenReport CURDIR & "\cr.RP1", "Weekly Sales Report", 0, "'", 6
Else
    OpenReport CURDIR & "\cr.RP1", "Ext WSR", 0, "'", 6
End If
End Sub
