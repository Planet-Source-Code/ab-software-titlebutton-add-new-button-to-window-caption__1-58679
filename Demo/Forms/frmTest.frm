VERSION 5.00
Object = "*\A..\..\absTitleButtons.vbp"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   3060
   StartUpPosition =   3  'Windows-Standard
   Begin absTitleButtons.TitleButton TitleButton1 
      Height          =   210
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   370
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub absTitleButton1_Click()
   Me.Hide
   frmTest2.Show
End Sub

Private Sub absTitleButton1_MouseHover()
   Me.Caption "Hover"
End Sub

Private Sub absTitleButton1_MouseLeave()
   Me.Caption "Leave"
End Sub

Private Sub Form_Load()
   Me.TitleButton1.Create Me.hWnd, "TrayIcon"
End Sub

Private Sub TitleButton1_Click()
   MsgBox "Button clicked!"
End Sub
