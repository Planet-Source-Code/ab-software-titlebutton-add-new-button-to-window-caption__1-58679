Attribute VB_Name = "modHook"
Option Explicit

'*********************
'* API Declarations  *
'*********************
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook&, ByVal lpfn&, ByVal hMod&, ByVal dwThreadId&) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook&) As Long

Private Type POINTAPI
   x As Long
   y As Long
End Type

'*********************
'* Type Declarations *
'*********************
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    Message As Long
    hWnd As Long
End Type

'*********************
'* Consts            *
'*********************
Private Const WM_MOVE = &H3             ' Window Message: Move window
Private Const WM_NCPAINT = &H85         ' Window Message: Repaint window
Private Const WM_SHOWWINDOW = &H18      ' Window Message: Show window
Private Const WM_ACTIVATEAPP = &H1C     ' Window Message: Activate application
Private Const WA_ACTIVE = 1             ' Window Activate: Activ
Private Const WA_INACTIVE = 0           ' Window Activate: Inactiv

Private Const SWP_FRAMECHANGED = &H20
Private Const GWL_WNDPROC = (-4)
Private Const GWL_EXSTYLE = -20

Private Const SM_CXFRAME = 32
Private Const SM_CYCAPTION = 4
Private Const SM_CXDLGFRAME = 7

Private Const GWL_STYLE = (-16)

'*********************
'* Vars              *
'*********************
Private WHook As Long

Public tButton As Object
Private m_hWndChild As Long
Private m_hWndParent As Long

Public Sub Init(ByVal hWndParent As Long, ByVal hWndChild As Long)
  
  m_hWndChild = hWndChild
  m_hWndParent = hWndParent
   
  'Initialize the window hooking for the button
  WHook = SetWindowsHookEx(4, AddressOf HookProc, 0, App.ThreadID)
   
  Call SetWindowLong(m_hWndChild, GWL_EXSTYLE, &H80)
  Call SetParent(m_hWndChild, GetParent(m_hWndParent))

End Sub

Public Sub Terminate()
  'Terminate the window hooking
  Call UnhookWindowsHookEx(WHook)

  Call SetParent(m_hWndChild, m_hWndParent)
End Sub

Public Function HookProc&(ByVal nCode&, ByVal wParam&, Inf As CWPSTRUCT)
   
'   On Error Resume Next
   
   Dim P As POINTAPI
   Dim ChildRect As RECT
   Dim FormRect As RECT
                    
   ' -----------------------------------------------------
   ' Check if cursor is over the title button
   ' -----------------------------------------------------
     
   ' Get the rectangle of the button
   FormRect = GetFormRect(m_hWndParent)
   ChildRect = GetButtonRect(FormRect)
   
   ' Get the cursor position
   GetCursorPos P
   
   ' Check if cursor is over the title button
   If P.x > ChildRect.Left And P.x < ChildRect.Left + ChildRect.Right Then
      If P.y > ChildRect.Top And P.y < ChildRect.Top + ChildRect.Bottom Then
         tButton.IsMouseOver = True
      Else
         tButton.IsMouseOver = False
     End If
   Else
      tButton.IsMouseOver = False
   End If
   
   ' -----------------------------------------------------
   ' Subclass the form
   ' -----------------------------------------------------
   If Inf.hWnd = m_hWndParent Then
      
      Select Case Inf.Message
               
         Case WM_NCPAINT, WM_MOVE:
              'Get the size of the Form
              FormRect = GetFormRect(m_hWndParent)
              'Place the button int the Titlebar
              ChildRect = GetButtonRect(FormRect)
              Call SetWindowPos(m_hWndChild, 0, ChildRect.Left, ChildRect.Top, ChildRect.Right, ChildRect.Bottom, SWP_FRAMECHANGED)
              Exit Function
               
         Case WM_ACTIVATEAPP
   
              If Inf.wParam = WA_ACTIVE Then
                 
                 'Our window is active
                 tButton.IsActive = True
                 
              ElseIf Inf.wParam = WA_INACTIVE Then
                 
                 'Another application's window has been activated
                 tButton.IsActive = False
                 
              End If
              
              Exit Function
              
         Case WM_SHOWWINDOW:
              
              If Inf.lParam = 0 And Inf.wParam = 0 Then
                 
                 tButton.IsMouseOver = False
                 SetParent m_hWndChild, m_hWndParent
                 
                 Exit Function
                 
              ElseIf Inf.lParam = 0 And Inf.wParam = 1 Then
                 
                 SetParent m_hWndChild, GetParent(m_hWndParent)
                 
              End If
              
              Exit Function
                          
      End Select

   End If
          
End Function

Private Function GetFormRect(ByVal hWnd As Long) As RECT
   
   GetWindowRect hWnd, GetFormRect
 
End Function

Private Function GetButtonRect(RForm As RECT) As RECT
   
   On Error Resume Next
   
   Dim l As Long
   Dim RChild As RECT
   Dim RTitle As RECT
   
   RTitle = WindowCaptionRect(m_hWndParent)
        
   If tButton.IsThemed Then
      l = RForm.Right - RTitle.Left
      l = l - RTitle.Left * 1.5
      l = l - RTitle.Bottom * 3

      With RChild
         .Top = RForm.Top + (RTitle.Top * 1.5)
         .Bottom = RTitle.Bottom - (RTitle.Top * 2)
         .Left = l
         .Right = RTitle.Bottom - (RTitle.Top * 2)
      End With
   Else
      l = RForm.Right - RTitle.Left
      l = l - RTitle.Left
      l = l - RTitle.Bottom * 3
   
      With RChild
         .Top = RForm.Top + (RTitle.Top * 1.5)
         .Bottom = RTitle.Bottom - (RTitle.Top * 2)
         .Left = l
         .Right = RTitle.Bottom - (RTitle.Top * 1.5)
      End With
   End If
   
   GetButtonRect = RChild
   
End Function

Private Function WindowCaptionRect(hWnd As Long) As RECT
   
   Dim r As RECT
   Dim XBorder As Long
   Dim fStyle As Long
   Dim YHeight As Long

   YHeight = GetSystemMetrics(SM_CYCAPTION)
   fStyle = GetWindowLong(hWnd, GWL_STYLE)
   
   Select Case fStyle And &H80
      Case &H80:       XBorder = GetSystemMetrics(SM_CXDLGFRAME)
      Case Else:       XBorder = GetSystemMetrics(SM_CXFRAME)
   End Select

   r.Left = XBorder
   r.Right = XBorder
   r.Top = XBorder
   r.Bottom = r.Top + YHeight - 1

   WindowCaptionRect = r

End Function
