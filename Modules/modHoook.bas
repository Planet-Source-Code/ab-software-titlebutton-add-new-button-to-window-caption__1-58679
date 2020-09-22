Attribute VB_Name = "modHook"
Option Explicit

' ************************************************************************
' Developer:  Robert E. Phelps
' Feel free to use this code as you wish, but please give credit to the
' author.  Do not sell this code; if you do, I want a piece of it :) !!!
' ************************************************************************

' Application defined Window properties
Public Const HWND_PROP_lpPrevWndProc = "lpPrevWndProc"
Public Const HWND_PROP_hWndMainForm = "hWndMainForm"
Public Const HWND_PROP_hWndTitleBarButton = "hWndTitleBarButton"
Public Const HWND_PROP_hWndTitleBarFakeButton = "hWndTitleBarFakeButton"
Public Const HWND_PROP_TitleBarButtonWidth = "TitleBarButtonWidth"

' Rectangle coordinates
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

' Windows messages
Public Const WM_SIZE = &H5
Public Const WM_LBUTTONUP = &H202           ' Used for command button
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_SETFONT = &H30

' Button messages
Public Const BM_SETCHECK = &HF1             ' Used for toggle button
Public Const BM_CLICK = &HF5                ' Used for command button

' GetWindowLong
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

' Window styles
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_THICKFRAME = &H40000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000

' Button styles
Public Const BS_PUSHBUTTON = &H0&           ' Used for command button
Public Const BS_AUTOCHECKBOX = &H3&         ' Used for toggle button
Public Const BS_PUSHLIKE = &H1000           ' Used for toggle button

' WindProc
Public Const GWL_WNDPROC As Long = -4&

' SetWindowPos
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10

' ShowWindow
Public Const SW_SHOWNOACTIVATE = 4

' SystemMetrics
Public Const SM_CXFIXEDFRAME = 7
Public Const SM_CXSIZEFRAME = 32
Public Const SM_CXSIZE = 30
Public Const SM_CXSMSIZE = 52
Public Const SM_CYCAPTION = 4
Public Const SM_CYSMCAPTION = 51

' Window position
Public Type WINDOWPOS
        hWnd As Long
        hWndInsertAfter As Long
        x As Long
        y As Long
        cx As Long
        cy As Long
        flags As Long
End Type

' Logical font
Public Const LF_FACESIZE = 32
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

' SystemParametersInfo
Public Const SPI_GETNONCLIENTMETRICS = 41
Public Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type

' API declarations
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Sub GetHILOWORD(lParam As Long, LOWORD As Long, HIWORD As Long)

    ' LOWORD of the lParam
    LOWORD = lParam And &HFFFF&
    ' LOWORD now equals 65,535 or &HFFFF

    ' HIWORD of the lParam
    HIWORD = lParam \ &H10000 And &HFFFF&
    ' HIWORD now equals 30,583 or &H7777

End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim udtRECT As RECT
    Dim udtWPOS As WINDOWPOS
    Dim x As Long
    Dim y As Long
    Dim lCapBtnWid As Long
    Dim lBrdWid As Long
    Dim lBtnTop As Long
    Dim lBtnLeft As Long
    Dim lBtnWidth As Long
    Dim lBtnHeight As Long

    Dim hWndMain As Long
    Dim hWndButton As Long
    Dim hWndFakeButton As Long
    Dim lpWndProc As Long


    lpWndProc = GetProp(hWnd, HWND_PROP_lpPrevWndProc)
    hWndMain = GetProp(hWnd, HWND_PROP_hWndMainForm)
    hWndButton = GetProp(hWnd, HWND_PROP_hWndTitleBarButton)

    Select Case hWnd
    Case hWndButton
        WindowProc = CallWindowProc(lpWndProc, hWnd, uMsg, wParam, lParam)
        Select Case uMsg
        Case WM_LBUTTONUP
            If BS_AUTOCHECKBOX <> (GetWindowLong(hWnd, GWL_STYLE) And BS_AUTOCHECKBOX) Then
                ' Command button
                GetHILOWORD lParam, x, y
                Call GetWindowRect(hWndButton, udtRECT)
                If x >= 0 And x <= udtRECT.Right - udtRECT.Left And _
                   y >= 0 And y <= udtRECT.Bottom - udtRECT.Top Then
                    ' Get the handle to control on form that currently has the focus
                    Call SetFocus(hWndMain)
                    hWndButton = GetFocus()

                    ' Send click message to fake button on form
                    hWndFakeButton = GetProp(hWnd, HWND_PROP_hWndTitleBarFakeButton)
                    Call SendMessage(hWndFakeButton, BM_CLICK, 0, 0)

                    ' Return focus to control on form that had the focus originally
                    Call SetFocus(hWndButton)
                End If
            End If
        Case BM_SETCHECK
            ' Toggle button
            hWndFakeButton = GetProp(hWnd, HWND_PROP_hWndTitleBarFakeButton)
            Call SendMessage(hWndFakeButton, BM_SETCHECK, wParam, 0)
            Call SetFocus(hWndMain)
        End Select
    Case hWndMain
        Select Case uMsg
        Case WM_WINDOWPOSCHANGED, WM_SIZE
            If uMsg = WM_WINDOWPOSCHANGED Then
                ' Get Main form RECT from WINDOWPOS passed in lParam
                CopyMemory udtWPOS, ByVal lParam, Len(udtWPOS)
                udtRECT.Left = udtWPOS.x
                udtRECT.Right = udtWPOS.x + udtWPOS.cx
                udtRECT.Top = udtWPOS.y
                udtRECT.Bottom = udtWPOS.y + udtWPOS.cy
            Else
                ' WM_SIZE, so get Main form RECT
                Call GetWindowRect(hWndMain, udtRECT)
            End If
            If WS_EX_TOOLWINDOW <> (GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_TOOLWINDOW) Then
                If WS_THICKFRAME <> (GetWindowLong(hWnd, GWL_STYLE) And WS_THICKFRAME) Then
                    ' Fixed-Single and Fixed-Dialog
                    lBrdWid = GetSystemMetrics(SM_CXFIXEDFRAME)
                    lCapBtnWid = GetSystemMetrics(SM_CXSIZE)
                    lBtnHeight = GetSystemMetrics(SM_CYCAPTION) - 5
                Else
                    ' Sizable
                    lBrdWid = GetSystemMetrics(SM_CXSIZEFRAME)
                    lCapBtnWid = GetSystemMetrics(SM_CXSIZE)
                    lBtnHeight = GetSystemMetrics(SM_CYCAPTION) - 5
                End If
            Else
                If WS_THICKFRAME <> (GetWindowLong(hWnd, GWL_STYLE) And WS_THICKFRAME) Then
                    ' Fixed ToolWindow
                    lBrdWid = GetSystemMetrics(SM_CXFIXEDFRAME)
                    lCapBtnWid = GetSystemMetrics(SM_CXSMSIZE)
                    lBtnHeight = GetSystemMetrics(SM_CYSMCAPTION) - 5
                Else
                    ' Sizable ToolWindow
                    lBrdWid = GetSystemMetrics(SM_CXSIZEFRAME)
                    lCapBtnWid = GetSystemMetrics(SM_CXSMSIZE)
                    lBtnHeight = GetSystemMetrics(SM_CYSMCAPTION) - 5
                End If
            End If

            ' Button pos
            lBtnTop = udtRECT.Top + lBrdWid + 2
            lBtnWidth = GetProp(hWndButton, HWND_PROP_TitleBarButtonWidth)
            If lBtnWidth = 0 Then
                ' Use default caption button width
                lBtnWidth = lCapBtnWid - 2
            End If
            If (WS_MINIMIZEBOX = (GetWindowLong(hWnd, GWL_STYLE) And WS_MINIMIZEBOX) Or _
               WS_MAXIMIZEBOX = (GetWindowLong(hWnd, GWL_STYLE) And WS_MAXIMIZEBOX)) And _
               WS_EX_TOOLWINDOW <> (GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_TOOLWINDOW) Then
                ' Calculate left pos of Min button
                lBtnLeft = udtRECT.Right - lBrdWid - lCapBtnWid - lCapBtnWid - lCapBtnWid + 2
                ' Calculate left pos of Titlebar button
                lBtnLeft = lBtnLeft - lBtnWidth - 2
            Else
                ' Calculate left pos of X button
                lBtnLeft = udtRECT.Right - lBrdWid - lCapBtnWid
                ' Calculate left pos of Titlebar button
                lBtnLeft = lBtnLeft - lBtnWidth - 2
            End If

            ' Position the button over the Titlebar of the form; remember, no parent (0)
            Call SetWindowPos(hWndButton, 0, lBtnLeft, lBtnTop, lBtnWidth, lBtnHeight, SWP_FRAMECHANGED + SWP_NOACTIVATE)
        End Select
        WindowProc = CallWindowProc(lpWndProc, hWnd, uMsg, wParam, lParam)
    End Select

End Function


