VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private DC As Long, DIB As Long
Private Pointer As Long
Private Bytes() As Byte

Private mWidth As Integer, mHeight As Integer

Private Type BITMAPINFOHEADER '40 bytes
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

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long





' DIBSection 的建立

Private Function CreateDIBSec(ByVal hDC As Long, ByVal Width As Integer, ByVal Height As Integer, Optional ByRef Pointer As Long = 0)
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
    CreateDIBSec = CreateDIBSection(hDC, BInfo, 0, VarPtr(Pointer), 0, 0)
End Function

Private Sub Class_Initialize()
    DC = CreateCompatibleDC(0)
    mWidth = 0
    mHeight = 0
End Sub

Private Sub Class_Terminate()
    DeleteDC DC
    DeleteObject DIB
End Sub

Public Sub CreateDIB(ByVal hDC As Long, ByVal Width As Integer, ByVal Height As Integer)
    Static BI As SAFEARRAY2D

    If Width = mWidth And Height = mHeight Then Exit Sub

    DIB = CreateDIBSec(hDC, Width, Height, Pointer)
    DeleteObject SelectObject(DC, DIB)
    
    With BI
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = 4 * (Width)
        .pvData = Pointer
    End With
    CopyMemory ByVal VarPtrArray(Bytes()), VarPtr(BI), 4
End Sub

Public Property Get hDC() As Long
    hDC = DC
End Property





' 點陣操作

Public Sub SetPoint(ByVal Color As Long, ByVal X As Integer, ByVal Y As Integer)
    Bytes(4 * X + 2, Y) = Color And &HFF
    Bytes(4 * X + 1, Y) = (Color \ &H100) And &HFF
    Bytes(4 * X, Y) = (Color \ &H10000) And &HFF
End Sub

Public Sub SetPointRGB(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal X As Integer, ByVal Y As Integer)
    Bytes(4 * X + 2, Y) = R
    Bytes(4 * X + 1, Y) = G
    Bytes(4 * X, Y) = B
End Sub

Public Function GetPoint(ByVal X As Integer, ByVal Y As Integer) As Long
    GetPoint = RGB(Bytes(4 * X + 2, Y), Bytes(4 * X + 1, Y), Bytes(4 * X, Y))
End Function

Public Function GetPointR(ByVal X As Integer, ByVal Y As Integer) As Byte
    GetPointR = Bytes(4 * X + 2, Y)
End Function

Public Function GetPointG(ByVal X As Integer, ByVal Y As Integer) As Byte
    GetPointG = Bytes(4 * X + 1, Y)
End Function

Public Function GetPointB(ByVal X As Integer, ByVal Y As Integer) As Byte
    GetPointB = Bytes(4 * X, Y)
End Function
