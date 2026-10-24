Attribute VB_Name = "modWin32"
Option Explicit

' Win32 API declarations used by the engine.

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long
End Type

Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Public Const SRCCOPY As Long = &HCC0020
Public Const DIB_RGB_COLORS As Long = 0
Public Const COLORONCOLOR As Long = 3
Public Const TRANSPARENT As Long = 1
Public Const TA_LEFT As Long = 0
Public Const TA_CENTER As Long = 6
Public Const FW_BOLD As Long = 700
Public Const GWL_WNDPROC As Long = -4
Public Const WM_SYSCOMMAND As Long = &H112&
Public Const SC_KEYMENU As Long = &HF100&

Private OldWndProc As Long     ' the form's original window procedure while hooked

Public Const VK_LBUTTON As Long = &H1&
Public Const VK_RETURN As Long = &HD&
Public Const VK_SHIFT As Long = &H10&
Public Const VK_CONTROL As Long = &H11&
Public Const VK_MENU As Long = &H12&
Public Const VK_ESCAPE As Long = &H1B&
Public Const VK_SPACE As Long = &H20&
Public Const VK_LEFT As Long = &H25&
Public Const VK_UP As Long = &H26&
Public Const VK_RIGHT As Long = &H27&
Public Const VK_DOWN As Long = &H28&
Public Const VK_1 As Long = &H31&
Public Const VK_2 As Long = &H32&
Public Const VK_3 As Long = &H33&
Public Const VK_4 As Long = &H34&
Public Const VK_A As Long = &H41&
Public Const VK_D As Long = &H44&
Public Const VK_E As Long = &H45&
Public Const VK_S As Long = &H53&
Public Const VK_W As Long = &H57&

Public Function IsKeyDown(ByVal vk As Long) As Boolean
    IsKeyDown = (GetAsyncKeyState(vk) And &H8000) <> 0
End Function

' Tapping Alt would put the window into menu mode, which stalls the game loop
' inside DoEvents. The form is subclassed so that request is dropped.

Public Sub HookWindow(ByVal hWnd As Long)
    OldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf GameWndProc)
End Sub

Public Sub UnhookWindow(ByVal hWnd As Long)
    If OldWndProc = 0 Then Exit Sub
    SetWindowLong hWnd, GWL_WNDPROC, OldWndProc
    OldWndProc = 0
End Sub

Public Function GameWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_SYSCOMMAND And (wParam And &HFFF0&) = SC_KEYMENU Then Exit Function
    GameWndProc = CallWindowProc(OldWndProc, hWnd, uMsg, wParam, lParam)
End Function
