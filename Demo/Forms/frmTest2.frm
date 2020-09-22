VERSION 5.00
Begin VB.Form frmTest2 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "FORM1"
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "END"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   Unload frmTest
   Unload frmTest2
   Set frmTest = Nothing
End Sub

Private Sub Command2_Click()
   frmTest.Show
   Unload Me
End Sub
