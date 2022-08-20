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

Private mintRed As Double
Private mintGreen As Double
Private mintBlue As Double

Private msngHue As Double
Private msngSaturation As Double
Private msngLuminosity As Double
Private OI As Double

Public Property Let Color(lColor As Long)
    mlColor = lColor
    mintRed = mlColor And &HFF
    mintGreen = (mlColor \ &H100) And &HFF
    mintBlue = (mlColor \ &H10000) And &HFF
    CalcHSL
End Property

Public Property Get Color() As Long
    Color = RGB(CInt(mintRed), CInt(mintGreen), CInt(mintBlue))
End Property

Property Let H(intHue As Double)
    msngHue = intHue
    CalcRGB
End Property

Property Get H() As Double
    H = msngHue
End Property

Property Let S(intSaturation As Double)
    msngSaturation = intSaturation / 100!
    CalcRGB
End Property

Property Get S() As Double
    S = msngSaturation * 100!
End Property

Property Let I(intLuminosity As Double)
    msngLuminosity = ValueInRange(intLuminosity / 100!, 0, 1)
    ChangeI
    OI = msngLuminosity
End Property

Property Get I() As Double
    I = msngLuminosity * 100!
End Property

Private Sub CalcHSL()
    Dim sngMx As Double
    Dim sngMn As Double
    Dim sngDelta As Double
    Dim sngPctRed As Double
    Dim sngPctGrn As Double
    Dim sngPctBlu As Double
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
    If msngLuminosity = 0 Or msngLuminosity = 1 Then msngHue = 0
    
    OI = msngLuminosity
End Sub

Private Sub CalcRGB()
    Dim sngM1 As Double
    Dim sngM2 As Double
    Dim sngPctRed As Double
    Dim sngPctGrn As Double
    Dim sngPctBlu As Double
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
    mintRed = 255 * sngPctRed
    mintGreen = 255 * sngPctGrn
    mintBlue = 255 * sngPctBlu
End Sub

Private Function rgbVal(sngN1 As Double, sngN2 As Double, sngHue As Double) As Double
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

Private Function sngMaxOf(sngV1 As Double, sngV2 As Double) As Double
    sngMaxOf = IIf(sngV1 > sngV2, sngV1, sngV2)
End Function

Private Function sngMinOf(sngV1 As Double, sngV2 As Double) As Double
   sngMinOf = IIf(sngV1 < sngV2, sngV1, sngV2)
End Function

Private Sub ChangeI()
    ' 只有當前後都在或都不在 0.5 以下時才有簡單換算
    If 0 < OI And OI < 0.5 And msngLuminosity < 0.5 Then
        mintRed = mintRed / OI * msngLuminosity
        mintGreen = mintGreen / OI * msngLuminosity
        mintBlue = mintBlue / OI * msngLuminosity
    ElseIf OI < 1 And OI > 0.5 And msngLuminosity > 0.5 Then
        mintRed = 255 - (255 - mintRed) / (1 - OI) * (1 - msngLuminosity)
        mintGreen = 255 - (255 - mintGreen) / (1 - OI) * (1 - msngLuminosity)
        mintBlue = 255 - (255 - mintBlue) / (1 - OI) * (1 - msngLuminosity)
    Else
        CalcRGB
    End If
End Sub

