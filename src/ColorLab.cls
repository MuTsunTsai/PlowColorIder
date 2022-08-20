VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

' 本模組根據 http://www.easyrgb.com/ 提供的公式計算

Private Const Gamma = 2.4

Private mvarCIEL As Double
Private mvarCIEa As Double
Private mvarCIEb As Double

Private mvarR As Single
Private mvarG As Single
Private mvarB As Single


Public Property Let Color(ByVal vData As Long)
    mvarR = vData And &HFF
    mvarG = (vData \ &H100) And &HFF
    mvarB = (vData \ &H10000) And &HFF
    CalcLab
End Property

Public Property Get Color() As Long
    Color = RGB(CInt(mvarR), CInt(mvarG), CInt(mvarB))
End Property

Public Property Let CIEb(ByVal vData As Double)
    mvarCIEb = vData
    CalcRGB
End Property

Public Property Get CIEb() As Double
    CIEb = RoundOff(mvarCIEb)
End Property

Public Property Let CIEa(ByVal vData As Double)
    mvarCIEa = vData
    CalcRGB
End Property

Public Property Get CIEa() As Double
    CIEa = RoundOff(mvarCIEa)
End Property

Public Property Let CIEL(ByVal vData As Double)
    mvarCIEL = vData
    CalcRGB
End Property

Public Property Get CIEL() As Double
    CIEL = RoundOff(mvarCIEL)
End Property



Private Sub CalcLab()
    Static var_R As Double, var_G As Double, var_B As Double
    Static var_X As Double, var_Y As Double, var_Z As Double

    var_R = mvarR / 255
    var_G = mvarG / 255
    var_B = mvarB / 255
    
    If var_R > 0.04045 Then var_R = ((var_R + 0.055) / 1.055) ^ Gamma Else var_R = var_R / 12.92
    If var_G > 0.04045 Then var_G = ((var_G + 0.055) / 1.055) ^ Gamma Else var_G = var_G / 12.92
    If var_B > 0.04045 Then var_B = ((var_B + 0.055) / 1.055) ^ Gamma Else var_B = var_B / 12.92
    
    var_X = var_R * 0.412453 + var_G * 0.35758 + var_B * 0.180423
    var_Y = var_R * 0.212671 + var_G * 0.71516 + var_B * 0.072169
    var_Z = var_R * 0.019334 + var_G * 0.119193 + var_B * 0.950227
         
    var_X = var_X * 100 / 95.0456
    var_Y = var_Y * 100 / 100#
    var_Z = var_Z * 100 / 108.8764
    
    If var_X > 0.00885645 Then var_X = var_X ^ (1 / 3) Else var_X = (7.78704 * var_X) + (16 / 116)
    If var_Y > 0.00885645 Then var_Y = var_Y ^ (1 / 3) Else var_Y = (7.78704 * var_Y) + (16 / 116)
    If var_Z > 0.00885645 Then var_Z = var_Z ^ (1 / 3) Else var_Z = (7.78704 * var_Z) + (16 / 116)

    mvarCIEL = (116 * var_Y) - 16
    mvarCIEa = 500 * (var_X - var_Y)
    mvarCIEb = 200 * (var_Y - var_Z)
End Sub

Private Sub CalcRGB()
    Static var_R As Double, var_G As Double, var_B As Double
    Static var_X As Double, var_Y As Double, var_Z As Double

    var_Y = (mvarCIEL + 16) / 116
    var_X = mvarCIEa / 500 + var_Y
    var_Z = var_Y - mvarCIEb / 200

    If var_Y ^ 3 > 0.00885645 Then var_Y = var_Y ^ 3 Else var_Y = (var_Y - 16 / 116) / 7.787
    If var_X ^ 3 > 0.00885645 Then var_X = var_X ^ 3 Else var_X = (var_X - 16 / 116) / 7.787
    If var_Z ^ 3 > 0.00885645 Then var_Z = var_Z ^ 3 Else var_Z = (var_Z - 16 / 116) / 7.787

    var_X = var_X * 95.0456 / 100
    var_Y = var_Y * 100# / 100
    var_Z = var_Z * 108.8764 / 100

    var_R = var_X * 3.240479 + var_Y * -1.53715 + var_Z * -0.498535
    var_G = var_X * -0.969256 + var_Y * 1.875992 + var_Z * 0.041556
    var_B = var_X * 0.055648 + var_Y * -0.204043 + var_Z * 1.057311

    If var_R > 0.0031308 Then var_R = 1.055 * (var_R ^ (1 / Gamma)) - 0.055 Else var_R = 12.92 * var_R
    If var_G > 0.0031308 Then var_G = 1.055 * (var_G ^ (1 / Gamma)) - 0.055 Else var_G = 12.92 * var_G
    If var_B > 0.0031308 Then var_B = 1.055 * (var_B ^ (1 / Gamma)) - 0.055 Else var_B = 12.92 * var_B

    mvarR = var_R * 255
    mvarG = var_G * 255
    mvarB = var_B * 255
End Sub

Private Function RoundOff(ByVal vData As Double) As Single
    RoundOff = CInt(vData * 100) / 100
End Function

