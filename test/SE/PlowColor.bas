Attribute VB_Name = "PlowColor"

Option Explicit

Public Const HC_ACTION = 0
Public Const WH_KEYBOARD = 2
Public Const GWL_WNDPROC = (-4)
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_ACTIVATE = &H6
Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2
Public Const WA_INACTIVE = 0

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type PointAPI
    X As Long
    Y As Long
End Type

Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public hnexthookproc As Long, PrevWndProc As Long
Private KBDHooked As Boolean

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static C As Object, fActive As Integer, T As Integer
    If uMsg = WM_ACTIVATE Then
        fActive = CInt(wParam And &HFFFF)
        If fActive = WA_INACTIVE Then
            UnHookKBD False
        Else
            If KBDHooked Then EnableKBDHook
        End If
    End If
    WndProc = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)
End Function

Public Sub UnHookKBD(Optional ByVal S As Boolean = True)
    If hnexthookproc <> 0 Then
        If S Then KBDHooked = False
        UnhookWindowsHookEx hnexthookproc
        hnexthookproc = 0
    End If
End Sub

Public Function EnableKBDHook()
    KBDHooked = True
    If hnexthookproc <> 0 Then Exit Function
    hnexthookproc = SetWindowsHookEx(WH_KEYBOARD, AddressOf MyKBHFunc, App.hInstance, 0)
    If hnexthookproc <> 0 Then EnableKBDHook = hnexthookproc
End Function

Public Function MyKBHFunc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    MyKBHFunc = 0
    If iCode < 0 Then MyKBHFunc = CallNextHookEx(hnexthookproc, iCode, wParam, lParam): Exit Function
    If (wParam >= vbKeyNumpad0 And wParam <= vbKeyNumpad9) _
        Or (wParam >= vbKey0 And wParam <= vbKey9) _
        Or (wParam >= vbKeyEnd And wParam <= vbKeyDown) _
        Or (GetAsyncKeyState(vbKeyControl) <> 0 And _
            (wParam = vbKeyC Or wParam = vbKeyV Or wParam = vbKeyA Or wParam = vbKeyZ)) _
        Or wParam = vbKeyDelete Or wParam = vbKeyBack _
        Or wParam = vbKeyTab Then
            Call CallNextHookEx(hnexthookproc, iCode, wParam, lParam)
    Else
        MyKBHFunc = 1
    End If
End Function

Public Sub LockCursor(ByVal ctlHwnd As Long, ByVal BorderPx As Integer)
    Static rect5 As RECT, res As Long
    GetWindowRect ctlHwnd, rect5
    rect5.Top = rect5.Top + BorderPx
    rect5.Left = rect5.Left + BorderPx
    rect5.Right = rect5.Right - BorderPx
    rect5.Bottom = rect5.Bottom - BorderPx
    res = ClipCursor(rect5)
End Sub

Public Sub UnLockCursor()
    Static rscreen As RECT
    rscreen.Top = 0
    rscreen.Left = 0
    rscreen.Right = Screen.Width / Screen.TwipsPerPixelX
    rscreen.Bottom = Screen.Height / Screen.TwipsPerPixelY
    ClipCursor rscreen
End Sub

