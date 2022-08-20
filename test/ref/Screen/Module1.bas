Attribute VB_Name = "Module1"

Public Const WH_KEYBOARD = 2
Public Const WH_KEYBOARD_LL = 13
Public Const WH_MOUSE = 7
Public Const WH_MOUSE_LL = 14
Public Const WH_JOURNALRECORD = 0
Public Const WH_JOURNALPLAYBACK = 1
Public Const WH_CALLWNDPROC = 4
Public Const WH_MSGFILTER = (-1)
Public Const WH_SYSMSGFILTER = 6

Public Const WM_ACTIVATE = &H6
Public Const WM_DROPFILES = &H233
Public Const WM_NOTIFY = &H4E
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSEWHEEL = &H20A

Public Const WM_MENUSELECT = &H11F


Public Type PointAPI
    X As Long
    Y As Long
End Type

Public Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As PointAPI
End Type

Public Type MSLLHookStruct
    pt As PointAPI
    mouseData As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Type EVENTMSG
        message As Long
        paramL As Long
        paramH As Long
        time As Long
        hwnd As Long
End Type

Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public MShnexthookproc As Long










Public Sub UnHookMS()
    Form1.Timer1.Enabled = False
    If MShnexthookproc <> 0 Then
        UnhookWindowsHookEx MShnexthookproc
        MShnexthookproc = 0
    End If
End Sub

Public Function EnableMSHook()
    If MShnexthookproc <> 0 Then Exit Function
    MShnexthookproc = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MyMSFunc, App.hInstance, 0)
    If MShnexthookproc <> 0 Then EnableMSHook = MShnexthookproc
End Function


Public Function MyMSFunc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static MSHook As MSLLHookStruct
    CopyMemory MSHook, ByVal lParam, Len(MSHook)
    
    If iCode < 0 Then
        MyMSFunc = CallNextHookEx(MShnexthookproc, iCode, wParam, lParam)
        Exit Function
    End If
    
    If wParam = WM_MOUSEMOVE Or wParam = WM_LBUTTONUP Or wParam = WM_LBUTTONDOWN Then
        MyMSFunc = 0
        Select Case wParam
            Case WM_MOUSEMOVE
                Form1.Scr MSHook.pt.X, MSHook.pt.Y
            Case WM_LBUTTONUP
                If Not ( _
                    MSHook.pt.X > Form1.Left / Screen.TwipsPerPixelX _
                    And MSHook.pt.X < (Form1.Left + Form1.Width) / Screen.TwipsPerPixelX _
                    And MSHook.pt.Y > Form1.Top / Screen.TwipsPerPixelY _
                    And MSHook.pt.Y < (Form1.Top + Form1.Height) / Screen.TwipsPerPixelY _
                    ) Then
                    
                    UnHookMS
                    MyMSFunc = 1
                End If
            Case WM_LBUTTONDOWN
                If Not ( _
                    MSHook.pt.X > Form1.Left / Screen.TwipsPerPixelX _
                    And MSHook.pt.X < (Form1.Left + Form1.Width) / Screen.TwipsPerPixelX _
                    And MSHook.pt.Y > Form1.Top / Screen.TwipsPerPixelY _
                    And MSHook.pt.Y < (Form1.Top + Form1.Height) / Screen.TwipsPerPixelY _
                    ) Then
                    MyMSFunc = 1
                End If
        End Select
    Else
        MyMSFunc = 1
    End If
    Call CallNextHookEx(MShnexthookproc, iCode, wParam, lParam)
End Function
