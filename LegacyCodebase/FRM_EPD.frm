VERSION 5.00
Begin VB.Form FRM_EPD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Product Detailes"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   Icon            =   "FRM_EPD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "&Products"
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Van"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   180
         Width           =   855
      End
   End
End
Attribute VB_Name = "FRM_EPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String

Private Sub Command1_Click()
If str = "Van" Then
    M_FRM_EVP.Show 1
Else
    M_FRM_HPE.Show 1
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Option1_Click()
str = "Van"
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    str = "Prd"
    Command1.SetFocus
End If
End Sub

Private Sub Option2_Click()
str = "Prd"
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    str = "Prd"
    Command1.SetFocus
End If
End Sub
