VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_MAT 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "POP Material and Equipments"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "M_FRM_MAT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
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
            Picture         =   "M_FRM_MAT.frx":0442
            Key             =   "p"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAT.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAT.frx":0D4E
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAT.frx":10A2
            Key             =   "e"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAT.frx":16DA
            Key             =   "s"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAT.frx":1A2E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAT.frx":1E8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAT.frx":2182
            Key             =   "pr"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAT.frx":258E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAT.frx":269A
            Key             =   "pre"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAT.frx":29EE
            Key             =   "brw"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_FRM_MAT.frx":2E02
            Key             =   "cpy"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5200
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   430
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   9155
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Category"
      TabPicture(0)   =   "M_FRM_MAT.frx":334E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text7"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "List2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "M_FRM_MAT.frx":336A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "list1"
      Tab(1).Control(2)=   "Label20"
      Tab(1).ControlCount=   3
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "M_FRM_MAT.frx":3386
         Left            =   1800
         List            =   "M_FRM_MAT.frx":33B4
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1770
         MaxLength       =   25
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3600
         Width           =   945
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3120
         Width           =   945
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1770
         MaxLength       =   25
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1785
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -74880
         TabIndex        =   10
         Top             =   4560
         Width           =   8295
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Material Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   480
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Material Short Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3480
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Value           =   1  'Checked
            Width           =   2415
         End
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1785
         MaxLength       =   25
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2745
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1785
         MaxLength       =   25
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2745
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   720
         Width           =   1065
      End
      Begin MSComctlLib.ListView list1 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   945
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Material Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Material ID"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Short Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1005
         TabIndex        =   19
         Top             =   3675
         Width           =   345
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   765
         TabIndex        =   17
         Top             =   3195
         Width           =   585
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   975
         TabIndex        =   15
         Top             =   2715
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   990
         TabIndex        =   13
         Top             =   2228
         Width           =   360
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Material"
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
         Left            =   -74640
         TabIndex        =   8
         Top             =   600
         Width           =   7905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Material ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   630
         TabIndex        =   6
         Top             =   795
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Material Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   330
         TabIndex        =   5
         Top             =   1275
         Width           =   1020
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Short Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   525
         TabIndex        =   4
         Top             =   1755
         Width           =   840
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
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
            Key             =   "p"
            Object.ToolTipText     =   "Print (Ctrl+P)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "pr"
            Object.ToolTipText     =   "Print Preview (Ctrl+R)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "h"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ed"
            Object.ToolTipText     =   "View (Ctrl+E)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s3"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pre"
            Object.ToolTipText     =   "Prefix Category"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "crt"
            Object.ToolTipText     =   "Create by (Ctrl+C)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "brw"
            Object.ToolTipText     =   "Browse (Ctrl+B)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
Attribute VB_Name = "M_FRM_MAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mod1 As Boolean, idnew As Boolean
Dim ls As ListItem

Private Sub CREATE()
mod1 = False
Text1 = ""
save
End Sub

Public Sub delete()
SYS_PSW.Caption = "Material Deletion"
SYS_PSW.Show 1
If M_FRM_MAT.Tag = "Delete" Then

    If MsgBox("Are you sure you want to delete Material ?", vbYesNo + vbCritical, "Effmcg") = vbYes Then
        If check("MAT", "MAT_IDY", Text1) = False Then
            MsgBox "Material Not Found, Define Material !", vbExclamation, "Effmcg"
            SSTab1.Tab = 0
            clrctr M_FRM_MAT
            Text1.Enabled = True
            Text1.SetFocus
            Exit Sub
        Else
            If check("prd", "MAT_IDY", Trim(Text1)) = False Then
                con.Execute "delete from MAT where MAT_IDY like '" & Sincrement(Text1) & "'"
                SSTab1.Tab = 0
                Text1.Enabled = True
                Text1.SetFocus
                clrctr M_FRM_MAT
            Else
                MsgBox "Cannot Delete, Transction Exists", vbExclamation, "Effmcg"
                Text2.SetFocus
                Exit Sub
            End If
        End If
        cancel
    Else
        cancel
    End If
Else
    Text2.SetFocus
End If
End Sub

Private Sub prdgrplist()
Dim RS As New ADODB.Recordset
If Check1.Value = 1 Then
    If RS.State = 1 Then RS.Close
    RS.Open "select * from MAT where mat_nme like '" & Text1 & "%'", con, adOpenKeyset, adLockOptimistic
Else
    If RS.State = 1 Then RS.Close
    RS.Open "select * from MAT where mat_sme like '" & Text1 & "%'", con, adOpenKeyset, adLockOptimistic
End If
LIST1.ListItems.Clear
If RS.RecordCount > 0 Then
    For A = 1 To RS.RecordCount
        Set ls = LIST1.ListItems.Add(, , Trim(RS("mat_nme")))
        ls.ListSubItems.Add , , Trim(RS("MAT_IDY"))
        ls.ListSubItems.Add , , Trim(RS("mat_sme"))
        RS.MoveNext
    Next
    SSTab1.Tab = 1
    LIST1.Visible = True
    LIST1.SetFocus
Else
    MsgBox "Material Not Found, Define Material !", vbExclamation, "Effmcg"
    SSTab1.Tab = 0
    Text1.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub save()
Dim RS As New ADODB.Recordset

If MsgBox("Confirm Record Save ?", vbExclamation + vbYesNo, "Effmcg") = vbNo Then
    cancel
    Exit Sub
End If

If Len(Trim(Text2)) = 0 Then
MsgBox "Material Should Not be Blank, Give Material !", vbExclamation, "Effmcg"
Text1.Enabled = False
SSTab1.Tab = 0
Text2.SetFocus
Exit Sub
End If

If RS.State = 1 Then RS.Close
If mod1 = True Then
    RS.Open "select * from MAT where MAT_IDY like '" & Trim(Text1) & "'", con, adOpenKeyset, adLockOptimistic
Else
    RS.Open "select * from MAT", con, adOpenKeyset, adLockOptimistic
    RS.AddNew
    RS("MAT_IDY") = increment("idy_mat")
    MsgBox "New Material ID is " & RS("MAT_IDY"), vbExclamation, "Effmcg"
End If
RS("mat_nme") = Trim(Text2)
RS("mat_sme") = Trim(Text3)
RS("mat_typ") = Trim(list2)
RS("mat_cod") = Trim(Text5)
RS("mat_qty") = Val(Text6)
RS("mat_rat") = Val(Text7)
RS.Update
clrctr M_FRM_MAT
Text1.Enabled = True
Text1.SetFocus
idnew = False
mod1 = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub cancel()
clrctr M_FRM_MAT
Text1.Enabled = True
SSTab1.Tab = 0
Text1.SetFocus
idnew = False
mod1 = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub newrec()
mod1 = False
SSTab1.Tab = 0
clrctr M_FRM_MAT
Text1.Enabled = False
Text2.SetFocus
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub exitfrm()
Unload Me
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
ElseIf Check1.Value = 0 Then
    Check2.Value = 1
End If
prdgrplist
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
ElseIf Check2.Value = 0 Then
    Check1.Value = 1
End If
prdgrplist
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
If Text1.Enabled = True Then
Text1.SetFocus
Else
Text2.SetFocus
End If
SSTab1.TabEnabled(1) = False
End Sub

Private Sub Form_KeyDown(k As Integer, s As Integer)
If k = 120 Then Shell "c:\windows\calc.exe"
If k > 111 And k < 120 Then If TypeOf M_FRM_MAT.ActiveControl Is TextBox Then M_FRM_MAT.ActiveControl.Text = Trim(FUNKEY(k))
If LIST1.Visible = False Then
    If k = 78 And s = 2 Then
        newrec
    ElseIf k = 83 And s = 2 Then
        save
    ElseIf k = 81 And s = 2 Then
        cancel
    ElseIf k = 88 And s = 2 Then
        exitfrm
    ElseIf k = 68 And s = 2 Then
        delete
    ElseIf k = 67 And s = 2 Then
        If Toolbar1.Buttons(12).Enabled = True Then CREATE
    ElseIf k = 69 And s = 2 Then
        If M_FRM_MAT.ActiveControl.name = "Text1" Then prdgrplist
    ElseIf k = 112 Then
        M_HELP
    ElseIf k = 80 And s = 2 Then
        Openreport CURDIR & "\lists.rp1", "Material list", 0, "", 6
    ElseIf k = 80 And s = 2 Then
        'MDIForm1.RRReport1.ReportName = "Categories list.RRW"
        'MDIForm1.RRReport1.Destination = 6
        'MDIForm1.RRReport1.RunReport 1
    End If
End If
End Sub

Private Sub Form_Keypress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If LIST1.Visible = True Then
        LIST1.Visible = False
        SSTab1.Tab = 0
        If check("MAT", "MAT_IDY", Trim(Text1)) = False Then
            Text1.SetFocus
        Else
            Text1.Enabled = False
            Text2.SetFocus
        End If
    Else
        If M_FRM_MAT.ActiveControl.name = "Text1" Then
            Unload Me
        ElseIf M_FRM_MAT.ActiveControl.name = "SSTab1" Then
        Else
            save
        End If
    End If
End If
End Sub

Private Sub Form_Load()
checkdongle
mod1 = False
idnew = False
Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub LIST1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If LIST1.SortOrder = lvwAscending Then
LIST1.SortOrder = lvwDescending
ElseIf LIST1.SortOrder = lvwDescending Then
LIST1.SortOrder = lvwAscending1
End If
LIST1.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub List1_DblClick()
list1_KeyDown 13, 0
End Sub

Private Sub list1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Press <Enter> or <Double Click> to Select Material or Press <Esc> to Quit"
End Sub

Private Sub list1_KeyDown(KeyAscii As Integer, s As Integer)
Dim RS As New ADODB.Recordset
If KeyAscii = 13 Then
    If LIST1.ListItems.Count > 0 Then
        If RS.State = 1 Then RS.Close
        Dim lst As ListSubItem
        Set lst = LIST1.SelectedItem.ListSubItems.Item(1)
        RS.Open "select * from MAT where MAT_IDY like '" & Trim(lst.Text) & "'", con, adOpenStatic
        If RS.RecordCount > 0 Then
            Text1 = Trim(RS("MAT_IDY"))
            Text2 = Trim(RS("mat_nme"))
            Text3 = Trim(RS("Mat_sme") & "")
            list2 = Trim(RS("mat_typ") & "")
            Text5 = Trim(RS("mat_cod") & "")
            Text6 = Trim(RS("mat_qty") & "")
            Text7 = Trim(RS("mat_rat") & "")
            LIST1.Visible = False
            SSTab1.Tab = 0
            Text1.Enabled = False
            Text2.SetFocus
            mod1 = True
            Toolbar1.Buttons(12).Enabled = True
        End If
    End If
End If
End Sub


Private Sub List2_Keydown(KeyAscii As Integer, s As Integer)
If KeyAscii = 13 Then
    Text5.SetFocus
ElseIf KeyAscii = 38 And s = 1 Then
    Text3.SetFocus
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Toolbar1.Enabled = True
LIST1.Visible = False
If SSTab1.Tab = 0 Then
    If Text1.Enabled = True Then
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Text1_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Type Material ID or Press <Enter> to Get List of Material ID's"
End Sub

Private Sub Text1_KeyPress(k As Integer)
If k > 96 And k < 123 Then k = k - 32
If k = 39 Or k = 41 Then k = 0
End Sub

Private Sub Text2_KeyPress(k As Integer)
If k > 96 And k < 123 Then k = k - 32
If k = 39 Or k = 41 Then k = 0
End Sub
Private Sub Text3_KeyPress(k As Integer)
If k > 96 And k < 123 Then k = k - 32
If k = 39 Or k = 41 Then k = 0
End Sub

Private Sub Text2_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Material Name"
End Sub

Private Sub Text3_GotFocus()
M_MDI.STS_BAR.Panels(1).Text = "Enter Short Name"
End Sub

Private Sub Text1_Keydown(KeyCode As Integer, Shift As Integer)
Dim RS As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
    If RS.State = 1 Then RS.Close
        RS.Open "select * from MAT where MAT_IDY like '" & Sincrement(Text1) & "'", con, adOpenStatic, adLockReadOnly
        If RS.RecordCount > 0 Then
        Text1 = Sincrement(Trim(Text1))
        Text2 = Trim(RS("Mat_nme"))
        Text3 = Trim(RS("Mat_sme"))
        list2 = Trim(RS("mat_typ"))
        Text5 = Val(RS("mat_qty"))
        Text6 = Val(RS("mat_rat"))
        Text7 = Val(RS("mat_cod"))
        Text1.Enabled = False
        Text2.SetFocus
        mod1 = True
        Toolbar1.Buttons(12).Enabled = True
        Else
            prdgrplist
        End If
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Text2)) > 0 Then
        Text3.SetFocus
    Else
        MsgBox "Material Should Not Be Blank, Give Material", vbExclamation, "Effmcg"
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RS As New ADODB.Recordset
If KeyCode = 13 Or KeyCode = 40 Then
    If Len(Trim(Text3)) = 0 Then
        MsgBox "Short Name Should Not Be Blank, Give Short Name", vbExclamation, "Effmcg"
        Text3.SetFocus
    Else
        list2.SetFocus
    End If
ElseIf KeyCode = 38 Then
    If Len(Trim(Text3)) > 0 Then
            Text2.SetFocus
    Else
        MsgBox "Short Name Should Not Be Blank, Give Short Name", vbExclamation, "Effmcg"
        Text3.SetFocus
    End If
End If
End Sub

Private Sub Text5_keydown(k As Integer, s As Integer)
If k = 13 Or k = 40 Then
    Text6.SetFocus
ElseIf k = 38 Then
    list2.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(k As Integer)
If k > 96 And k < 123 Then k = k - 32
If k = 39 Or k = 41 Then k = 0
End Sub

Private Sub Text6_keydown(k As Integer, s As Integer)
If k = 13 Or k = 40 Then
    Text7.SetFocus
ElseIf k = 38 Then
    Text5.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
End Sub

Private Sub Text7_KeyDown(k As Integer, s As Integer)
If k = 13 Or k = 40 Then
    save
ElseIf k = 38 Then
    Text6.SetFocus
End If
End Sub

Private Sub Text7_keypress(k As Integer)
If k > 47 And k < 58 Or k = 8 Or k = 46 Then
Else
k = 0
End If
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
ElseIf Button.KEY = "ed" Then
    If M_FRM_MAT.ActiveControl.name = "Text1" Then prdgrplist
ElseIf Button.KEY = "d" Then
    delete
ElseIf Button.KEY = "pre" Then
    M_FRM_PRE.Caption = "Material ID Prefix"
    M_FRM_PRE.Show 1
ElseIf Button.KEY = "p" Then
    Openreport CURDIR & "\lists.rp1", "Materials List", 0, "", 6
ElseIf Button.KEY = "pr" Then
    Exit Sub
er1:
    MsgBox "Material Not Found, Define Material !", vbExclamation, "Effmcg"
ElseIf Button.KEY = "crt" Then
    CREATE
End If
End Sub
