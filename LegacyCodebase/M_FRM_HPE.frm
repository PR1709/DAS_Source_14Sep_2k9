VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FRM_HPE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HHT Products Export"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   Icon            =   "M_FRM_HPE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton Command3 
         Caption         =   "&Select All"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   5640
         TabIndex        =   1
         Top             =   4560
         Width           =   2295
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   1200
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&OK"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
      End
      Begin MSComctlLib.ListView LVMOD 
         Height          =   4125
         Left            =   240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   7276
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
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
            Text            =   "Mobile No."
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Activation Status"
            Object.Width           =   3351
         EndProperty
      End
   End
End
Attribute VB_Name = "M_FRM_HPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If str <> "" Then
    Dim FS_SOU As Object
    Set FS_SOU = CreateObject("Scripting.filesystemobject")
    If FS_SOU.folderexists(App.Path & "\HHT") = False Then
    FS_SOU.CREATEFOLDER (App.Path & "\HHT")
    End If
    
    If FS_SOU.FILEExists(App.Path & "\HHT\Retailer.txt") = True Then
    FS_SOU.deleteFILE (App.Path & "\HHT\Retailer.txt")
    End If
    
    Dim hht_rst1 As New ADODB.Recordset
    
    hht_rst1.Open "select * from prd", CON, adOpenKeyset, adLockOptimistic
        
    Open App.Path & "\HHT\Products.txt" For Output As #1   'Open file for output.
    While hht_rst1.EOF = False
    Write #1, hht_rst1(0) & "#" & hht_rst1(1)
    hht_rst1.MoveNext
    Wend
    Close #1
    If hht_rst1.State = 1 Then hht_rst1.Close
    str = ""
    MsgBox "Retailers data exported to HHT", vbInformation + vbYesNo, "Effmcg"
Else
    MsgBox "Selection is not made!", vbInformation + vbYesNo, "Effmcg"
    optsal.SetFocus
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "&Select All" Then
    For I = 1 To LVMOD.ListItems.Count
        LVMOD.ListItems(I).Checked = True
    Next
    Command3.Caption = "&Deselect All"
Else
    For I = 1 To LVMOD.ListItems.Count
        LVMOD.ListItems(I).Checked = False
    Next
    Command3.Caption = "&Select All"
End If
End Sub

Private Sub Form_Activate()
'If lbltyp.Caption <> "Van" Then
    
    Dim HHT_RST As New ADODB.Recordset
    HHT_RST.Open "select * from prd,brd where prd.prd_idy like brd.prd_idy order by prd.prd_idy", CON, adOpenStatic
    If HHT_RST.EOF = False Then
        LVMOD.Visible = True
        LVMOD.ColumnHeaders.Clear
        LVMOD.ColumnHeaders.Add , , "Product Name", 2500
        LVMOD.ColumnHeaders.Add , , "Rate (B)", 1200
        LVMOD.ColumnHeaders.Add , , "Batch Name", 750
        LVMOD.ColumnHeaders.Add , , "Date of Mfg", 1200
        LVMOD.ColumnHeaders.Add , , "Date of Exp", 1250
        LVMOD.ColumnHeaders.Add , , "Prd.Id", 1500
        LVMOD.ColumnHeaders.Add , , "Bat.Id", 1500
        
        LVMOD.ListItems.Clear
        While Not HHT_RST.EOF
        If Val(HHT_RST("PRD_QTY")) > 0 Then
            Set ls = LVMOD.ListItems.Add(, , Trim(HHT_RST("prd_NME")) & "")
            ls.ListSubItems.Add , , convert(HHT_RST("SAL_RAT")) & "" 'QTY(RS("saf_lev") & "")
            ls.ListSubItems.Add , , Trim(HHT_RST("bat_nme")) & ""
            ls.ListSubItems.Add , , Trim(HHT_RST("BRD_DOM")) & ""
            ls.ListSubItems.Add , , Trim(HHT_RST("Brd_DOE")) & ""
            ls.ListSubItems.Add , , Trim(HHT_RST("prd_IDY")) & ""
            ls.ListSubItems.Add , , Trim(HHT_RST("BAT_IDY")) & ""
                        
        End If
            HHT_RST.MoveNext
        Wend
    HHT_RST.Close
    Else
        MsgBox "Products not Found! Define Products", vbInformation + vbYesNo, "Effmcg"
        Unload Me
    End If
'End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
