VERSION 5.00
Begin VB.Form M_FRM_PRC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Rate Check"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   Icon            =   "M_FRM_PRC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "M_FRM_PRC.frx":0442
      Left            =   2400
      List            =   "M_FRM_PRC.frx":044C
      TabIndex        =   1
      Top             =   307
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enable Purchase Rate Check"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2115
   End
End
Attribute VB_Name = "M_FRM_PRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(K As Integer)
If K = 27 Then Unload Me
End Sub

Private Sub Form_Load()
If check("dir", "dir_idy like 'SYS' AND DIR_IDX LIKE 'PSW' AND DIR_MSG", "Y") = True Then
    list1.Text = "Y"
Else
    list1.Text = "N"
End If
End Sub

Private Sub List1_keypress(K As Integer)
If K = 13 Then
    If list1.Text = "Y" Then
        CON.Execute "update dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'PSW' SET DIR_MSG = 'Y'"
    Else
        CON.Execute "update dir where dir_idy like 'SYS' AND DIR_IDX LIKE 'PSW' SET DIR_MSG = ''"
    End If
    Unload Me
End If
End Sub
