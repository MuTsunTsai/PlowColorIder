VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private SR As Integer, SG As Integer, SB As Integer
Private R As String, G As String, B As String, SRGB As String, SBGR As String
Private NColor As Long
Private SColor As Long
Private CurrentLab As New ColorLab
Private staName As String, extName As String

Private mColor As Long

Public Function getGPer() As String
    getGPer = Mid(Str((SG / 2.55)), 2, 4) & "%"
End Function

Public Function getBPer() As String
    getBPer = Mid(Str((SB / 2.55)), 2, 4) & "%"
End Function

Public Function getRPer() As String
    getRPer = Mid(Str((SR / 2.55)), 2, 4) & "%"
End Function

Public Function getRGB() As String
    getRGB = "µL"
    If Mid(R, 1, 1) = Mid(R, 2, 1) And Mid(G, 1, 1) = Mid(G, 2, 1) And Mid(B, 1, 1) = Mid(B, 2, 1) Then getRGB = Hex(SR / 17) & Hex(SG / 17) & Hex(SB / 17)
End Function

Public Function getNRGB() As String
    Static tR As Integer, tG As Integer, tB As Integer
    tR = NColor And &HF
    tG = (NColor \ &H100) And &HF
    tB = (NColor \ &H10000) And &HF
    getNRGB = Hex(tR) & Hex(tG) & Hex(tB)
End Function

Public Function getNColor() As Long
    getNColor = NColor
End Function

Public Function getSRGB() As String
    Static tR As Integer, tG As Integer, tB As Integer
    tR = SColor And &HF
    tG = (SColor \ &H100) And &HF
    tB = (SColor \ &H10000) And &HF
    getSRGB = Hex(tR) & Hex(tG) & Hex(tB)
End Function

Public Function getSColor() As Long
    getSColor = SColor
End Function

Public Function getRRGGBB() As String
    getRRGGBB = SRGB
End Function

Public Function getBBGGRR() As String
    getBBGGRR = SBGR
End Function

Public Function getExtName() As String
    getExtName = extName
End Function

Public Function getStaName() As String
    getStaName = staName
End Function

Public Function FindName(ByVal n As String) As Long
    Static I As Integer
    FindName = -1
    For I = 1 To 140
        If YColor(I, 0) = n Then FindName = YColor(I, 1)
    Next I
End Function

Public Property Let Color(ByVal vData As Long)
    Static Color As Long, I As Integer
        
    mColor = vData
    SR = mColor And &HFF: SG = (mColor \ &H100) And &HFF: SB = (mColor \ &H10000) And &HFF
    R = Hex(SR): G = Hex(SG): B = Hex(SB)
    R = Right("0" & R, 2): G = Right("0" & G, 2): B = Right("0" & B, 2)
    
    CurrentLab.Color = mColor
    NColor = getNearestColor(17)
    SColor = getNearestColor(51)
    
    SRGB = R & G & B
    SBGR = B & G & R
    
    staName = "µL"
    For I = 1 To 16
        If dColor(I, 1) = mColor Then staName = dColor(I, 0)
    Next I
    
    extName = "µL"
    For I = 1 To 140
        If YColor(I, 1) = mColor Then extName = YColor(I, 0)
    Next I
    
End Property

Public Property Get Color() As Long
    Color = mColor
End Property





Private Function Distance(ByVal dColor As Long) As Double
    Static Lab As New ColorLab
    Lab.Color = dColor
    'CIE 76
    Distance = ( _
            (CurrentLab.CIEL - Lab.CIEL) ^ 2 + _
            (CurrentLab.CIEa - Lab.CIEa) ^ 2 + _
            (CurrentLab.CIEb - Lab.CIEb) ^ 2 _
        ) ^ 0.5
        
    'CIE 94
    'Static C1 As Double, C2 As Double
    'Static H1 As Double, H2 As Double
    'C1 = (CurrentLab.CIEa ^ 2 + CurrentLab.CIEb ^ 2) ^ 0.5
    'C2 = (Lab.CIEa ^ 2 + Lab.CIEb ^ 2) ^ 0.5
    'If CurrentLab.CIEa <> 0 Then H1 = Atn(CurrentLab.CIEb / CurrentLab.CIEa) Else H1 = 0
    'If Lab.CIEa <> 0 Then H2 = Atn(Lab.CIEb / Lab.CIEa) Else H2 = 0
    'Distance = ( _
            (CurrentLab.CIEL - Lab.CIEL) ^ 2 + _
            ((C1 - C2) / (1 + 0.045 * C2)) ^ 2 + _
            ((H1 - H2) / (1 + 0.015 * C2)) ^ 2 _
        ) ^ 0.5
End Function

Private Function getNearestColor(ByVal Unit As Integer) As Long
    Static tR As Integer, tG As Integer, tB As Integer
    Static Dis1 As Double, Dis2 As Double
    tR = Int(SR / Unit) * Unit
    tG = Int(SG / Unit) * Unit
    tB = Int(SB / Unit) * Unit
    
    getNearestColor = RGB(tR, tG, tB)
    Dis1 = Distance(RGB(tR, tG, tB))
    Dis2 = Distance(RGB(tR, tG, tB + Unit))
    If Dis2 < Dis1 Then getNearestColor = RGB(tR, tG, tB + Unit): Dis1 = Dis2
    Dis2 = Distance(RGB(tR, tG + Unit, tB))
    If Dis2 < Dis1 Then getNearestColor = RGB(tR, tG + Unit, tB): Dis1 = Dis2
    Dis2 = Distance(RGB(tR, tG + Unit, tB + Unit))
    If Dis2 < Dis1 Then getNearestColor = RGB(tR, tG + Unit, tB + Unit): Dis1 = Dis2
    Dis2 = Distance(RGB(tR + Unit, tG, tB))
    If Dis2 < Dis1 Then getNearestColor = RGB(tR + Unit, tG, tB): Dis1 = Dis2
    Dis2 = Distance(RGB(tR + Unit, tG, tB + Unit))
    If Dis2 < Dis1 Then getNearestColor = RGB(tR + Unit, tG, tB + Unit): Dis1 = Dis2
    Dis2 = Distance(RGB(tR + Unit, tG + Unit, tB))
    If Dis2 < Dis1 Then getNearestColor = RGB(tR + Unit, tG + Unit, tB): Dis1 = Dis2
    Dis2 = Distance(RGB(tR + Unit, tG + Unit, tB + Unit))
    If Dis2 < Dis1 Then getNearestColor = RGB(tR + Unit, tG + Unit, tB + Unit): Dis1 = Dis2
End Function
