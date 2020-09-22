VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl TitleButton 
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   Picture         =   "absTitleButton.ctx":0000
   ScaleHeight     =   2550
   ScaleWidth      =   4230
   Begin VB.PictureBox tbButton 
      Height          =   615
      Left            =   360
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin PicClip.PictureClip picClip 
      Left            =   240
      Top             =   1440
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   4
   End
   Begin PicClip.PictureClip picClassic 
      Left            =   240
      Top             =   2280
      _ExtentX        =   1693
      _ExtentY        =   370
      _Version        =   393216
      Cols            =   4
      Picture         =   "absTitleButton.ctx":0582
   End
End
Attribute VB_Name = "TitleButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ---------------------------------------------------------------
' Declarations
' ---------------------------------------------------------------

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" _
   Alias "RegOpenKeyExA" ( _
   ByVal HKey As Long, _
   ByVal lpSubKey As String, _
   ByVal ulOptions As Long, _
   ByVal samDesired As Long, _
   ByRef phkResult As Long _
) As Long

Private Declare Function RegQueryValueEx Lib "advapi32" _
   Alias "RegQueryValueExA" ( _
   ByVal HKey As Long, _
   ByVal lpValueName As String, _
   ByVal lpReserved As Long, _
   ByRef lpType As Long, _
   ByVal lpData As String, _
   ByRef lpcbData As Long _
) As Long


Private Declare Function RegCloseKey Lib "advapi32" ( _
   ByVal HKey As Long _
) As Long

' ---------------------------------------------------------------
' Constants
' ---------------------------------------------------------------

Private Const ICON_INACTIVE As Long = 0  ' Inactive icon
Private Const ICON_NORMAL As Long = 1    ' Normal icon
Private Const ICON_HOVER As Long = 2     ' Hover icon
Private Const ICON_MOUSEDOWN As Long = 3 ' Icon on mouse down

' Registry Root Keys
Private Const HKEY_CLASSES_ROOT As Long = &H80000000     ' Root
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002    ' Local Machine
Private Const HKEY_USERS As Long = &H80000003            ' Users
Private Const HKEY_CURRENT_USER As Long = &H80000001     ' Current User
Private Const HKEY_CURRENT_CONFIG As Long = &H80000005   ' Current Config

Private Const KEY_READ As Long = &H20019    ' Read Access
Private Const REG_SZ As Long = 1            ' VBNullChar terminated String


Private Const VER_PLATFORM_WIN32_NT As Long = 2

Private m_IsActive As Boolean
Private m_IsMouseDown As Boolean
Private m_IsMouseOver As Boolean
Private m_IsThemed As Boolean

Private Type OSVERSIONINFO
   OSVSize As Long
   dwVerMajor As Long
   dwVerMinor As Long
   dwBuildNumber As Long
   PlatformID As Long
   szCSDVersion As String * 128
End Type
' Registry keys
Public Enum VbHKey
   VbHKEY_CLASSES_ROOT = HKEY_CLASSES_ROOT
   VbHKEY_LOCAL_MACHINE = HKEY_LOCAL_MACHINE
   VbHKEY_USERS = HKEY_USERS
   VbHKEY_CURRENT_USER = HKEY_CURRENT_USER
   VbHKEY_CURRENT_CONFIG = HKEY_CURRENT_CONFIG
End Enum

' Windows scheme constants.
Public Enum VbWindowsScheme
   VbClassic = 0        ' Classic
   VbNormalColor = 1    ' Normal Color (Blue)
   VbMetallic = 2       ' Metallic (Silver)
   VbHomeStead = 3      ' HomeStead (Olive)
End Enum

Private m_hWndParent As Long

Event Click()
Event MouseHover()
Event MouseLeave()

' Creates the button and draws it in the title bar.
'
' @hWnd Window handle of the form.
Public Sub Create(ByVal hWndParent As Long, Optional ToolTipText As String)
     
   m_hWndParent = hWndParent
   
   With tbButton
   
      .ToolTipText = ToolTipText
           
      modHook.Init hWndParent, .hWnd
      
   End With
   
End Sub

Public Sub Terminate()
     
   Set tButton = Nothing
   
   modHook.Terminate
   
End Sub

Public Function Redraw()
   
   Dim Scheme As String
   Dim pic As StdPicture
   Dim icon As Long
   
   Scheme = GetSchemeStyle()
   
   ' ---------------------------------------------------
   ' Get the pictures from resource file
   ' ---------------------------------------------------
   
   If IsThemeAble And StrComp(Scheme, "@themeui.dll,-883", vbTextCompare) <> 0 Then
      
      On Error Resume Next
      Set pic = LoadResPicture(UCase$(GetSchemeStyle), vbResBitmap)
      On Error GoTo 0
      
      If pic Is Nothing Then
         Set pic = LoadResPicture("CLASSIC", vbResBitmap)
         Me.IsThemed = False
      Else
         Me.IsThemed = True
      End If
   Else
      Set pic = LoadResPicture("CLASSIC", vbResBitmap)
      Me.IsThemed = False
   End If
   
   picClip.Picture = pic
   
   ' ---------------------------------------------------
   ' Select the icon to display
   ' ---------------------------------------------------
   If IsActive Then
      icon = ICON_NORMAL
   Else
      icon = ICON_INACTIVE
   End If
      
   If m_IsMouseOver Then
      If IsMouseDown Then
         icon = ICON_MOUSEDOWN
      Else
         icon = ICON_HOVER
      End If
   Else
      If IsActive Then
         icon = ICON_NORMAL
      Else
         icon = ICON_INACTIVE
      End If
   End If
   If picClip.Picture <> 0 Then
      tbButton.Picture = picClip.GraphicCell(icon)
   End If
   
End Function

Private Sub tbButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   
   If Button = vbLeftButton Then
      If IsMouseOver Then
         IsMouseDown = True
      End If
   End If
   
   UserControl.Parent.SetFocus
   
End Sub

Private Sub tbButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton And IsMouseOver Then
      RaiseEvent Click
   End If
   IsMouseDown = False
End Sub


' Initialize the control by drawing the button.
'
Private Sub UserControl_Initialize()
   
   Dim wdh As Long
   Dim hgt As Long
   Dim rgn As Long

   ' Icon is normal.
   m_IsActive = True
   m_IsMouseDown = False
   m_IsMouseOver = False
   Me.Redraw
   
   UserControl.Picture = picClip.GraphicCell(1)
   
   ' Round the corners of the button.
   If Me.IsThemed Then
      
      ' Get the width and the height of the button.
      wdh = (tbButton.Width \ (Screen.TwipsPerPixelX * 2)) + 2
      hgt = (tbButton.Height \ (Screen.TwipsPerPixelY * 2)) + 2
      
      ' Create the reign for the round rectangle
      rgn = CreateRoundRectRgn(0, 0, wdh, hgt, 2, 2)
      
      ' Round the controls corners.
      SetWindowRgn tbButton.hWnd, rgn, True
      
   End If

   Set tButton = Me
   
End Sub

' Never resize the control.
'
Private Sub UserControl_Resize()
   
   On Error Resume Next
   
   UserControl.Height = picClip.CellHeight * Screen.TwipsPerPixelY
   UserControl.Width = picClip.CellWidth * Screen.TwipsPerPixelX

End Sub

Public Property Get IsActive() As Boolean
   IsActive = m_IsActive
End Property
Public Property Let IsActive(ByVal New_IsActive As Boolean)
   If m_IsActive <> New_IsActive Then
      m_IsActive = New_IsActive
      Redraw
   End If
End Property

Public Property Get IsMouseOver() As Boolean
   IsMouseOver = m_IsMouseOver
End Property
Public Property Let IsMouseOver(ByVal New_IsMouseOver As Boolean)
   
   If m_IsMouseOver <> New_IsMouseOver Then
      m_IsMouseOver = New_IsMouseOver
      Redraw
   End If
   
   If New_IsMouseOver Then
      RaiseEvent MouseHover
   Else
      RaiseEvent MouseLeave
   End If
   
End Property

Public Property Get IsMouseDown() As Boolean
   IsMouseDown = m_IsMouseDown
End Property

' Set <b>True</b> if a mouse button is pressed.
'
' @New_IsThemed Set <b>True</b> if a theme is active; <b>False</b> otherwise.
Public Property Let IsMouseDown(ByVal New_IsMouseDown As Boolean)
   If m_IsMouseDown <> New_IsMouseDown Then
      m_IsMouseDown = New_IsMouseDown
      Redraw
   End If
End Property

' This property returns <b>True</b> if a theme is active.
'
' @IsThemed Returns <b>True</b> if a theme is active; <b>False</b> otherwise.
Public Property Get IsThemed() As Boolean
   IsThemed = m_IsThemed
End Property

' Set <b>True</b> if a theme is active.
'
' @New_IsThemed Set <b>True</b> if a theme is active; <b>False</b> otherwise.
Public Property Let IsThemed(ByVal New_IsThemed As Boolean)
   m_IsThemed = New_IsThemed
End Property

Private Sub UserControl_Terminate()
   Me.Terminate
End Sub

' Returns <b>True</b> if operating system is Windows XP.
'
' @IsThemeAble Returns <b>True</b> if operating system is Windows XP; <b>False</b> otherwise.
Private Function IsThemeAble() As Boolean

   Dim OSV As OSVERSIONINFO
   
   OSV.OSVSize = Len(OSV)
   GetVersionEx OSV

   If OSV.PlatformID = VER_PLATFORM_WIN32_NT And _
      OSV.dwVerMinor = 1 Then
      IsThemeAble = True ' Windows XP
   Else
      IsThemeAble = False
   End If

End Function

' Returns the style of the selected theme scheme.
'
' @Scheme A constant of the VbWindowScheme enumeration.
Public Function GetSchemeStyle() As String
   
   On Error GoTo Error_Handle
   
   Dim SchemeName As String
   Dim RegKeyTheme As String
   
   RegKeyTheme = "Software\Microsoft\Windows\CurrentVersion\ThemeManager"
   
   SchemeName = GetKeyValue(VbHKEY_CURRENT_USER, RegKeyTheme, "ColorName")
   
   If Len(SchemeName) > 0 Then
      GetSchemeStyle = UCase$(SchemeName)
      Exit Function
   End If
   
   'Scheme = VbHomeStead
Finally:
      
   Exit Function
   
Error_Handle:

   GetSchemeStyle = "classic"
   
   GoTo Finally

End Function

Public Function GetKeyValue(ByVal MainKey As VbHKey, ByVal SubKey As String, ByVal Value As String) As String
   
   Dim RetVal As Long
   Dim HKey As Long
   Dim TmpSNum As String * 255
   
   RetVal = RegOpenKeyEx(MainKey, SubKey, 0&, KEY_READ, HKey)
   
   If RetVal <> 0 Then
      GetKeyValue = "Can't open the registry."
      Exit Function
   End If
   
   RetVal = RegQueryValueEx(HKey, Value, 0, REG_SZ, ByVal TmpSNum, Len(TmpSNum))
   
   If RetVal <> 0 Then
      GetKeyValue = "Can't read or find the registry."
      Exit Function
   End If
   
   GetKeyValue = Left$(TmpSNum, InStr(1, TmpSNum, vbNullChar) - 1)
   
   RetVal = RegCloseKey(HKey)
   
End Function

