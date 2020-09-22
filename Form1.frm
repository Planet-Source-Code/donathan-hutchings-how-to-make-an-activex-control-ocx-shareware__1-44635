VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   420
      Left            =   105
      TabIndex        =   0
      Top             =   3765
      Width           =   1290
   End
   Begin VB.Label lblDescription 
      Caption         =   $"Form1.frx":0000
      Height          =   1110
      Left            =   30
      TabIndex        =   1
      Top             =   90
      Width           =   5070
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Register
End Sub
