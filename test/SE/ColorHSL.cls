VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorHSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private mlColor As Long

Private mintRed As Integer
Private mintGreen As Integer
Private mintBlue As Integer

Private msngHue As Single
Private msngSaturation As Single
Private msngLuminosity As Single

Private mintCalc As Integer
Private Const RGB2HSL = 1
Private Const HSL2RGB = 2

Public Property Let Color(lColor As Long)
    mlColor = lColor
    mintRed = mlColor And &HFF
    mintGreen = (mlColor \ &H100) And &HFF
    mintBlue = (mlColor \ &H10000) And &HFF
    CalcHSL
End Property

Public Property Get Color() As Long
    CalcRGB
    Color = RGB(mintRed, mintGreen, mintBlue)
End Property

Property Let H(intHue As Single)
    msngHue = intHue
    mintCalc = HSL2RGB
End Property

Property Get H() As Single
    If mintCalc = RGB2HSL Then CalcHSL
    H = msngHue
End Property

Property Let S(intSaturation As Single)
    msngSaturation = intSaturation / 100!
    mintCalc = HSL2RGB
End Property

Property Get S() As Single
    If mintCalc = RGB2HSL Then CalcHSL
    S = msngSaturation * 100!
End Property

Property Let I(intLuminosity As Single)
    msngLuminosity = intLuminosity / 100!
    If msngLuminosity < 0 Then msngLuminosity = 0
    If msngLuminosity > 1 Then msngLuminosity = 1
    mintCalc = HSL2RGB
End Property

Property Get I() As Single
    If mintCalc = RGB2HSL Then CalcHSL
    I = msngLuminosity * 100!
End Property

Private Sub CalcHSL()
    Dim sngMx As Single
    Dim sngMn As Single
    Dim sngDelta As Single
    Dim sngPctRed As Single
    Dim sngPctGrn As Single
    Dim sngPctBlu As Single
    sngPctRed = mintRed / 255
    sngPctGrn = mintGreen / 255
    sngPctBlu = mintBlue / 255
    sngMx = sngMaxOf(sngMaxOf(sngPctRed, sngPctGrn), sngPctBlu)
    sngMn = sngMinOf(sngMinOf(sngPctRed, sngPctGrn), sngPctBlu)
    sngDelta = sngMx - sngMn
    msngLuminosity = (sngMx + sngMn) / 2
    If sngMx = sngMn Then msngSaturation = 0 Else msngSaturation = 1
    If msngLuminosity <= 0.5 Then
        If msngSaturation > 0 Then msngSaturation = sngDelta / (sngMx + sngMn)
    Else
        If msngSaturation > 0 Then msngSaturation = sngDelta / (2 - sngMx - sngMn)
    End If
    If msngSaturation Then
        If sngPctRed = sngMx Then msngHue = (sngPctGrn - sngPctBlu) / sngDelta
        If sngPctGrn = sngMx Then msngHue = 2 + (sngPctBlu - sngPctRed) / sngDelta
        If sngPctBlu = sngMx Then msngHue = 4 + (sngPctRed - sngPctGrn) / sngDelta
        msngHue = msngHue * 60
    End If
    If msngHue < 0 Then msngHue = msngHue + 360
    mintCalc = 0
End Sub

Private Sub CalcRGB()
    Dim sngM1 As Single
    Dim sngM2 As Single
    Dim sngPctRed As Single
    Dim sngPctGrn As Single
    Dim sngPctBlu As Single
    If msngSaturation = 0! Then
        sngPctRed = msngLuminosity
        sngPctGrn = msngLuminosity
        sngPctBlu = msngLuminosity
    Else
        If msngLuminosity <= 0.5 Then
            sngM2 = msngLuminosity * (1! + msngSaturation)
        Else
            sngM2 = msngLuminosity + msngSaturation - msngLuminosity * msngSaturation
        End If
        sngM1 = 2! * msngLuminosity - sngM2
        sngPctRed = rgbVal(sngM1, sngM2, msngHue + 120!)
        sngPctGrn = rgbVal(sngM1, sngM2, msngHue)
        sngPctBlu = rgbVal(sngM1, sngM2, msngHue - 120!)
    End If
    mintRed = Int(255.9999 * sngPctRed)
    mintGreen = Int(255.9999 * sngPctGrn)
    mintBlue = Int(255.9999 * sngPctBlu)
    mintCalc = 0
End Sub

Private Function rgbVal(sngN1 As Single, sngN2 As Single, sngHue As Single) As Single
   If sngHue > 360 Then
      sngHue = sngHue - 360
   ElseIf sngHue < 0 Then
      sngHue = sngHue + 360
   End If
   If sngHue < 60 Then
      rgbVal = sngN1 + (sngN2 - sngN1) * sngHue / 60
   ElseIf sngHue < 180 Then
      rgbVal = sngN2
   ElseIf sngHue < 240 Then
      rgbVal = sngN1 + (sngN2 - sngN1) * (240 - sngHue) / 60
   Else
      rgbVal = sngN1
   End If
End Function

Private Function sngMaxOf(sngV1 As Single, sngV2 As Single) As Single
    sngMaxOf = IIf(sngV1 > sngV2, sngV1, sngV2)
End Function

Private Function sngMinOf(sngV1 As Single, sngV2 As Single) As Single
   sngMinOf = IIf(sngV1 < sngV2, sngV1, sngV2)
End Function
