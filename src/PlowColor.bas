Attribute VB_Name = "PlowColor"

Option Explicit

Public Const MAX_PATH = 260
Public Const MAX_FILE = 260

Public Const MF_BYPOSITION = &H400&
Public Const MF_REMOVE = &H1000&

Public Const HC_ACTION = 0
Public Const GWL_WNDPROC = (-4)
Public Const GWL_STYLE = (-16)

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40

Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2
Public Const WA_INACTIVE = 0

Public Const WH_KEYBOARD = 2
Public Const WH_KEYBOARD_LL = 13
Public Const WH_MOUSE = 7
Public Const WH_MOUSE_LL = 14
Public Const WH_JOURNALRECORD = 0
Public Const WH_JOURNALPLAYBACK = 1

Public Const WS_CAPTION = &HC00000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000

Public Const WM_ACTIVATE = &H6
Public Const WM_ACTIVATEAPP = &H1C
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

Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2

Public Const GRADIENT_FILL_RECT_H = &H0
Public Const GRADIENT_FILL_RECT_V = &H1
Public Const GRADIENT_FILL_TRIANGLE = &H2

Private Const ICC_USEREX_CLASSES = &H200

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

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

Public Type KBDLLHookStruct
    vkCode As Long
    scanCode As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Type MSLLHookStruct
    pt As PointAPI
    mouseData As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Type BITMAPINFOHEADER '40 bytes
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

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Public Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Public Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Public Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Declare Function DragAcceptFiles Lib "shell32" (ByVal hWnd As Long, ByVal fAccept As Long) As Long
Public Declare Function DragFinish Lib "shell32" (ByVal hDrop As Long) As Long
Public Declare Function DragQueryFile Lib "shell32" Alias "DragQueryFileW" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As Long, ByVal ch As Long) As Long

Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Public KBDhnexthookproc As Long, PrevWndProc As Long
Public HSVMode As Boolean, DecMode As Boolean
Public dColor(16, 1), YColor(140, 1), Sys(27) As String

Public KBDHooked As Boolean
Public KBDOptionFocused As Integer






Public Sub Main()
    InitCommonControlsVB
    ColorInitialize
    frmMain.Show
End Sub

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function








Public Function getAppVersion(Optional ByVal full As Boolean = False) As String
    If full Then getAppVersion = "5.0.14" Else getAppVersion = "5.0"
    'getAppVersion = App.Major & "." & App.Minor & "." & App.Revision
    'getAppVersion = App.Major & "." & App.Minor
End Function








Public Function CreateDIBSec(ByVal hDC As Long, ByVal Width As Integer, ByVal Height As Integer)
    Static BInfo As BITMAPINFO
    BInfo.bmiHeader.biSize = 40
    BInfo.bmiHeader.biWidth = Width
    BInfo.bmiHeader.biHeight = Height
    BInfo.bmiHeader.biPlanes = 1
    BInfo.bmiHeader.biBitCount = 32
    BInfo.bmiHeader.biCompression = 0
    BInfo.bmiHeader.biSizeImage = CLng(Width) * CLng(Height)
    BInfo.bmiHeader.biXPelsPerMeter = 0
    BInfo.bmiHeader.biYPelsPerMeter = 0
    BInfo.bmiHeader.biClrUsed = 0
    BInfo.bmiHeader.biClrImportant = 0
    CreateDIBSec = CreateDIBSection(hDC, BInfo, 0, 0, 0, 0)
End Function

Public Function ValueInRange(ByVal Value As Variant, ByVal Min As Variant, ByVal Max As Variant) As Variant
    ValueInRange = Value
    If Value < Min Then ValueInRange = Min
    If Value > Max Then ValueInRange = Max
End Function

Public Function CreateRect(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer) As RECT
    CreateRect.Left = Left: CreateRect.Top = Top
    CreateRect.Right = Right: CreateRect.Bottom = Bottom
End Function

Public Function Ptr2StrU(ByVal pAddr As Long) As String
    Static lenAddr As Long
    lenAddr = lstrlenW(pAddr)
    Ptr2StrU = Space$(lenAddr)
    CopyMemory ByVal StrPtr(Ptr2StrU), ByVal pAddr, lenAddr * 2
End Function





Public Function Color16(ByVal Clr As Long) As Integer
    Dim UINT As Long
    UINT = Clr * &H100&
    If UINT < &H7FFF Then
        Color16 = CInt(UINT)
    Else
        Color16 = CInt(UINT - &H10000)
    End If
End Function

Public Sub SetTriVertexColor(ByRef V As TRIVERTEX, ByVal C As Long)
    V.Red = Color16(C And &HFF)
    V.Green = Color16((C \ &H100) And &HFF)
    V.Blue = Color16((C \ &H10000) And &HFF)
End Sub




Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static fActive As Integer, I As Integer
    Static Handle&, FileNameBuff$, result&, hDrop&, fileNum&, fLine$
    
    Select Case uMsg
        Case WM_MOUSEWHEEL
            frmMain.Form_MouseWheel wParam
        Case WM_ACTIVATEAPP
            If wParam Then
                frmMain.Timer.Enabled = True
                EnableKBDHook
            Else
                If frmMain.ScrCapturing Then frmMain.ScrCaptureEnd
                frmMain.Timer.Enabled = False
                EnableSystemMenu frmMain.hWnd, True
                UnHookKBD
            End If
        Case WM_DROPFILES
            hDrop = wParam
            FileNameBuff = Space(MAX_PATH)
            ' Unicode API 使用通則：
            ' 將宣告的字串改成 Long，傳送字串時用 StrPtr 傳送位置
            ' 接收時用如下的程式碼處理
            DragQueryFile hDrop, 0, StrPtr(FileNameBuff), MAX_PATH
            frmMain.Form_FileDragDrop Trim(Ptr2StrU(StrPtr(FileNameBuff)))
            DragFinish hDrop
    End Select
    WndProc = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)
End Function

' 捲軸捲動公用程式
Public Sub ScrollBarScroll(ByRef sBar As Object, ByVal Speed As Integer)
    sBar.Value = ValueInRange(sBar.Value + Speed, 0, sBar.Max)
End Sub

' 自動觸發滑鼠移動公用程式
Public Sub RaiseMouseMove()
    Static P As PointAPI
    GetCursorPos P
    SetCursorPos P.X, P.Y
End Sub

' 開關視窗之系統選單
Public Sub EnableSystemMenu(ByVal hWnd As Long, ByVal bRevert As Boolean)
    Static hSysMenu As Long, nCnt As Long, I As Long
    If bRevert Then
        GetSystemMenu hWnd, True
        DrawMenuBar hWnd
    Else
        hSysMenu = GetSystemMenu(hWnd, False)
        If hSysMenu Then
            nCnt = GetMenuItemCount(hSysMenu)
            If nCnt Then
                For I = 0 To nCnt - 1
                    RemoveMenu hSysMenu, 0, MF_BYPOSITION Or MF_REMOVE
                Next I
            End If
        End If
    End If
End Sub










Public Sub UnHookKBD()
    If KBDhnexthookproc <> 0 Then
        UnhookWindowsHookEx KBDhnexthookproc
        KBDhnexthookproc = 0
    End If
End Sub

Public Function EnableKBDHook()
    If KBDhnexthookproc <> 0 Then Exit Function
    KBDhnexthookproc = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf MyKBDFunc, App.hInstance, 0)
    If KBDhnexthookproc <> 0 Then EnableKBDHook = KBDhnexthookproc
End Function

Public Function MyKBDFunc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static KBDHook As KBDLLHookStruct
    Static Shift As Integer
    CopyMemory KBDHook, ByVal lParam, Len(KBDHook)
    
    MyKBDFunc = 0
    
    If iCode < 0 Then
        MyKBDFunc = CallNextHookEx(KBDhnexthookproc, iCode, wParam, lParam)
        Exit Function
    End If
    
    ' 螢幕撿色時除了 Esc 以外停止一切按鍵的使用
    If frmMain.ScrCapturing Then
        If KBDHook.vkCode = vbKeyEscape Then frmMain.ScrCaptureEnd
        MyKBDFunc = 1
    End If
    
    ' 文字框用鎖定機制
    If KBDHooked And Not ( _
        (KBDHook.vkCode >= vbKeyNumpad0 And KBDHook.vkCode <= vbKeyNumpad9) _
        Or (KBDHook.vkCode >= vbKey0 And KBDHook.vkCode <= vbKey9) _
        Or (frmMain.Fra(0).Visible And Not DecMode And _
            KBDHook.vkCode >= vbKeyA And KBDHook.vkCode <= vbKeyF) _
        Or (KBDHook.vkCode >= vbKeyEnd And KBDHook.vkCode <= vbKeyDown) _
        Or ((GetAsyncKeyState(vbKeyControl) <> 0 Or wParam = WM_KEYUP) And _
            (KBDHook.vkCode = vbKeyC Or KBDHook.vkCode = vbKeyX Or KBDHook.vkCode = vbKeyV Or _
            KBDHook.vkCode = vbKeyA Or KBDHook.vkCode = vbKeyZ Or _
            KBDHook.vkCode = vbKeyT Or KBDHook.vkCode = vbKeyS)) _
        Or KBDHook.vkCode = vbKeyDelete Or KBDHook.vkCode = vbKeyBack _
        Or KBDHook.vkCode = vbKeyPageUp Or KBDHook.vkCode = vbKeyPageDown _
        Or KBDHook.vkCode = vbKeyTab _
        Or KBDHook.vkCode = 44 Or KBDHook.vkCode = 115 _
        Or KBDHook.vkCode = 91 Or KBDHook.vkCode = 92 Or KBDHook.vkCode = 93 _
        Or KBDHook.vkCode = 160 Or KBDHook.vkCode = 161 _
        Or KBDHook.vkCode = 162 Or KBDHook.vkCode = 163 _
        Or KBDHook.vkCode = 164 Or KBDHook.vkCode = 165 _
    ) Then MyKBDFunc = 1
    
    ' 91,92=Windows 鍵、93=Context 鍵
    ' 160~165=Control、Alt、Shift 鍵
    ' 44=Prt Sc、115=F4
    
    If KBDOptionFocused <> 0 And wParam = WM_KEYDOWN Then
        If KBDOptionFocused < 3 Then
            frmMain.Hand.OptionKeyDown KBDHook.vkCode
            If KBDHook.vkCode <> vbKeyTab And KBDHook.vkCode <> 160 And KBDHook.vkCode <> 161 Then MyKBDFunc = 1
        Else
            frmMain.HSColor.SetFocus
        End If
    End If
    
    ' 快速鍵處理
    Shift = 0
    If GetAsyncKeyState(vbKeyShift) <> 0 Then Shift = Shift + vbShiftMask
    If GetAsyncKeyState(vbKeyControl) <> 0 Then Shift = Shift + vbCtrlMask
    If GetAsyncKeyState(vbKeyMenu) <> 0 Then Shift = Shift + vbAltMask
    
    If wParam = WM_KEYDOWN And MyKBDFunc = 0 Then frmMain.Form_GlobalKeyDown KBDHook.vkCode, Shift
          
    Call CallNextHookEx(KBDhnexthookproc, iCode, wParam, lParam)
End Function








Public Sub LockCursor(ByVal ctlHwnd As Long, ByVal BorderPx As Integer)
    Static ctlRect As RECT, res As Long
    GetWindowRect ctlHwnd, ctlRect
    ctlRect.Top = ctlRect.Top + BorderPx
    ctlRect.Left = ctlRect.Left + BorderPx
    ctlRect.Right = ctlRect.Right - BorderPx
    ctlRect.Bottom = ctlRect.Bottom - BorderPx
    res = ClipCursor(ctlRect)
End Sub

Public Sub UnLockCursor()
    ClipCursor CreateRect(0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)
End Sub









Public Sub ColorInitialize()
    Dim Temp, I As Integer
    
    Temp = Array( _
        "aqua", &HFFFF00, "black", &H0&, "blue", &HFF0000, "fuchsia", &HFF00FF, _
        "gray", &H808080, "green", &H8000&, "lime", &HFF00&, "maroon", &H80&, _
        "navy", &H800000, "olive", &H8080&, "purple", &H800080, "red", &HFF&, _
        "silver", &HC0C0C0, "teal", &H808000, "white", &HFFFFFF, "yellow", &HFFFF& _
    )
    For I = 1 To 16
        dColor(I, 0) = Temp(I * 2 - 2)
        dColor(I, 1) = Temp(I * 2 - 1)
    Next I
    
    Temp = Array( _
        "aliceblue", &HFFF8F0, "antiquewhite", &HD7EBFA, "aqua", &HFFFF00, "quamarine", &HD4FF7F, "azure", &HFFFFF0, "beige", &HDCF5F5, "bisque", &HC4E4FF, "black", &H0&, "blanchedalmond", &HCDEBFF, "blue", &HFF0000, _
        "blueviolet", &HE22B8A, "brown", &H2A2AA5, "burlywood", &H87B8DE, "cadetblue", &HA09E5F, "chartreuse", &HFF7F&, "chocolate", &H1E69D2, "coral", &H507FFF, "cornflowerblue", &HED9564, "cornsilk", &HDCF8FF, "crimson", &H3C14DC, _
        "cyan", &HFFFF00, "darkblue", &H8B0000, "darkcyan", &H8B8B00, "darkgoldenrod", &HB86B8, "darkgray", &HA9A9A9, "darkgreen", &H6400&, "darkkhaki", &H6BB7BD, "darkmagenta", &H8B008B, "darkolivegreen", &H2F6B55, "darkorange", &H8CFF&, _
        "darkorchid", &HCC3299, "darkred", &H8B&, "darksalmon", &H7A96E9, "darkseagreen", &H8BBC8F, "darkslateblue", &H8B3D48, "darkslategray", &H4F4F2F, "darkturquoise", &HD1CE00, "darkviolet", &HD30094, "yellowgreen", &H32CD9A, "deeppink", &H9314FF, _
        "deepskyblue", &HFFBF00, "dimgray", &H696969, "dodgerblue", &HFF901E, "firebrick", &H2222B2, "floralwhite", &HF0FAFF, "forestgreen", &H228B22, "fuchsia", &HFF00FF, "gainsboro", &HDCDCDC, "ghostwhite", &HFFF8F8, "gold", &HD7FF&, _
        "goldenrod", &H20A5DA, "gray", &H808080, "green", &H8000&, "greenyellow", &H2FFFAD, "honeydew", &HF0FFF0, "hotpink", &HB469FF, "indianred", &H5C5CCD, "indigo", &H82004B, "ivory", &HF0FFFF, "khaki", &H8CE6F0, _
        "lavender", &HFAE6E6, "lavenderblush", &HF5F0FF, "lawngreen", &HFC7C&, "lemonchiffon", &HCDFAFF, "lightblue", &HE6D8AD, "lightcoral", &H8080F0, "lightcyan", &HFFFFE0, "lightgoldenrodyellow", &HD2FAFA, _
        "lightgreen", &H90EE90, "lightgrey", &HD3D3D3, "lightpink", &HC1B6FF, "lightsalmon", &H7AA0FF, "lightseagreen", &HAAB220, "lightskyblue", &HFACE87, "lightslategray", &H998877, "lightsteelblue", &HDEC4B0, _
        "lightyellow", &HE0FFFF, "lime", &HFF00&, "mediumseagreen", &H71B33C, "mediumslateblue", &HEE687B, "mediumspringgreen", &H9AFA00, "mediumturquoise", &HCCD148, "mediumvioletred", &H8515C7, "midnightblue", &H701919, _
        "mintcream", &HFAFFF5, "mistyrose", &HE1E4FF, "moccasin", &HB5E4FF, "navajowhite", &HADDEFF, "navy", &H800000, "oldlace", &HE6F5FD, "olive", &H8080&, "olivedrab", &H238E6B, "orange", &HA5FF&, "orangered", &H45FF&, _
        "orchid", &HD670DA, "palegoldenrod", &HAAE8EE, "palegreen", &H98FB98, "paleturquoise", &HEEEEAF, "palevioletred", &H9370DB, "papayawhip", &HD5EFFF, "peachpuff", &HB9DAFF, "peru", &H3F85CD, "pink", &HCBC0FF, "plum", &HDDA0DD, _
        "powderblue", &HE6E0B0, "purple", &H800080, "red", &HFF&, "rosybrown", &H8F8FBC, "royalblue", &HE16941, "saddlebrown", &H13458B, "salmon", &H7280FA, "sandybrown", &H60A4F4, "seagreen", &H578B2E, "seashell", &HEEF5FF, _
        "sienna", &H2D52A0, "silver", &HC0C0C0, "skyblue", &HEBCE87, "slateblue", &HCD5A6A, "slategray", &H908070, "snow", &HFAFAFF, "springgreen", &H7FFF00, "steelblue", &HB48246, "tan", &H8CB4D2, "teal", &H808000, _
        "thistle", &HD8BFD8, "tomato", &H4763FF, "turquoise", &HD0E040, "violet", &HEE82EE, "wheat", &HB3DEF5, "white", &HFFFFFF, "whitesmoke", &HF5F5F5, "yellow", &HFFFF&, "limegreen", &H32CD32, "linen", &HE6F0FA, _
        "magenta", &HFF00FF, "maroon", &H80&, "mediumaquamarine", &HAACD66, "mediumblue", &HCD0000, "mediumorchid", &HD355BA, "mediumpurple", &HDB7093 _
    )
    For I = 1 To 140
        YColor(I, 0) = Temp(I * 2 - 2)
        YColor(I, 1) = Temp(I * 2 - 1)
    Next I
    
    Temp = Array( _
        "activeborder", "activecaption", "appworkspace", "background", "buttonface", _
        "buttonhighlight", "buttonshadow", "buttontext", "captiontext", "graytext", _
        "highlight", "highlighttext", "inactiveborder", "inactivecaption", "inactivecaptiontext", _
        "infobackground", "infotext", "menu", "menutext", "scrollbar", "threeddarkshadow", _
        "threedface", "threedhighlight", "threedlightshadow", "threedshadow", "window", _
        "windowframe", "windowtext" _
    )
    For I = 0 To 27
        Sys(I) = Temp(I)
    Next I

End Sub

Public Function StringToColor(ByVal SC As String) As Long
    Static C As New ColorInfo
    If Left(SC, 1) = "#" Then SC = Right(SC, Len(SC) - 1)
    If Len(SC) = 6 And IsNumeric("&H" & SC) Then
        SC = Mid(SC, 5, 1) & Mid(SC, 6, 1) & Mid(SC, 3, 1) & Mid(SC, 4, 1) & Mid(SC, 1, 1) & Mid(SC, 2, 1)
        SC = Val("&H" & SC): If SC < 0 Then SC = SC + 65536
        StringToColor = SC
    ElseIf Len(SC) = 3 And IsNumeric("&H" & SC) Then
        SC = Mid(SC, 3, 1) & Mid(SC, 3, 1) & Mid(SC, 2, 1) & Mid(SC, 2, 1) & Mid(SC, 1, 1) & Mid(SC, 1, 1)
        SC = Val("&H" & SC): If SC < 0 Then SC = SC + 65536
        StringToColor = SC
    ElseIf C.FindName(SC) <> -1 Then
        StringToColor = C.FindName(SC)
    Else
        StringToColor = -1
    End If
End Function
