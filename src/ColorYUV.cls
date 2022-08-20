VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorYUV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarY As Single
Private mvarU As Single
Private mvarV As Single
Private mvarR As Single
Private mvarG As Single
Private mvarB As Single

Public Property Let V(ByVal vData As Single)
    mvarV = vData / 100
    CalcRGB
End Property

Public Property Get V() As Single
    V = 100 * RoundOff(mvarV)
End Property

Public Property Let U(ByVal vData As Single)
    mvarU = vData / 100
    CalcRGB
End Property

Public Property Get U() As Single
    U = 100 * RoundOff(mvarU)
End Property

Public Property Let Y(ByVal vData As Single)
    mvarY = vData / 100
    CalcRGB
End Property

Public Property Get Y() As Single
    Y = 100 * RoundOff(mvarY)
End Property

Public Property Let Color(ByVal vData As Long)
    mvarR = vData And &HFF
    mvarG = (vData \ &H100) And &HFF
    mvarB = (vData \ &H10000) And &HFF
    CalcYUV
End Property

Public Property Get Color() As Long
    Color = RGB(CInt(mvarR), CInt(mvarG), CInt(mvarB))
End Property

Private Sub CalcYUV()
    Static SR As Single, SG As Single, SB As Single
    SR = mvarR / 255: SG = mvarG / 255: SB = mvarB / 255
    mvarY = 0.299 * SR + 0.587 * SG + 0.114 * SB
    mvarU = -0.147 * SR - 0.289 * SG + 0.436 * SB
    mvarV = 0.615 * SR - 0.515 * SG - 0.1 * SB
End Sub

Private Sub CalcRGB()
    Static SR As Single, SG As Single, SB As Single
    SR = mvarY + 1.13983 * mvarV
    SG = mvarY - 0.39465 * mvarU - 0.5806 * mvarV
    SB = mvarY + 2.03211 * mvarU
    mvarR = 255 * SR
    mvarG = 255 * SG
    mvarB = 255 * SB
End Sub

Private Function RoundOff(ByVal vData As Double) As Single
    RoundOff = CInt(vData * 100) / 100
End Function



