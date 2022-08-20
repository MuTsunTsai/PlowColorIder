Attribute VB_Name = "Module1"

Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

'Public Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

'Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
'Public Declare Function SetDCBrushColor Lib "gdi32" (ByVal hDC As Long, ByVal colorref As Long) As Long

'Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
'Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

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

Public Const GRADIENT_FILL_RECT_H = &H0
Public Const GRADIENT_FILL_RECT_V = &H1
Public Const GRADIENT_FILL_TRIANGLE = &H2

Public Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Public Function ValueInRange(ByVal Value As Variant, ByVal Min As Variant, ByVal Max As Variant) As Variant
    ValueInRange = Value
    If Value < Min Then ValueInRange = Min
    If Value > Max Then ValueInRange = Max
End Function

Public Function CreateRect(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer) As RECT
    CreateRect.Left = Left: CreateRect.Top = Top
    CreateRect.Right = Right: CreateRect.Bottom = Bottom
End Function


