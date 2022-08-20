VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorCMYK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarColor As Long
Private mvarC As Single
Private mvarM As Single
Private mvarY As Single
Private mvarK As Single
Private mvarR As Integer
Private mvarG As Integer
Private mvarB As Integer

Public Property Let K(ByVal vData As Single)
    Static SR As Single, SG As Single, SB As Single
    mvarK = vData / 100
    SR = 1 - mvarC * (1 - mvarK) - mvarK
    SG = 1 - mvarM * (1 - mvarK) - mvarK
    SB = 1 - mvarY * (1 - mvarK) - mvarK
    mvarR = Int(255.9999 * SR)
    mvarG = Int(255.9999 * SG)
    mvarB = Int(255.9999 * SB)
    mvarColor = RGB(mvarR, mvarG, mvarB)
End Property

Public Property Get K() As Single
    K = mvarK * 100
End Property

Public Property Let Y(ByVal vData As Single)
    Static SB As Single
    mvarY = vData / 100
    SB = 1 - mvarY * (1 - mvarK) - mvarK
    mvarB = Int(255.9999 * SB)
    mvarColor = RGB(mvarR, mvarG, mvarB)
End Property

Public Property Get Y() As Single
    Y = mvarY * 100
End Property

Public Property Let M(ByVal vData As Single)
    Static SG As Single
    mvarM = vData / 100
    SG = 1 - mvarM * (1 - mvarK) - mvarK
    mvarG = Int(255.9999 * SG)
    mvarColor = RGB(mvarR, mvarG, mvarB)
End Property

Public Property Get M() As Single
    M = mvarM * 100
End Property

Public Property Let C(ByVal vData As Single)
    Static SR As Single
    mvarC = vData / 100
    SR = 1 - mvarC * (1 - mvarK) - mvarK
    mvarR = Int(255.9999 * SR)
    mvarColor = RGB(mvarR, mvarG, mvarB)
End Property

Public Property Get C() As Single
    C = mvarC * 100
End Property

Public Property Let Color(ByVal vData As Long)
    mvarColor = vData
    mvarR = mvarColor And &HFF
    mvarG = (mvarColor \ &H100) And &HFF
    mvarB = (mvarColor \ &H10000) And &HFF
    CalcCMYK
End Property

Public Property Get Color() As Long
    Color = mvarColor
End Property

Private Sub CalcCMYK()
    mvarC = 1 - mvarR / 255: mvarM = 1 - mvarG / 255: mvarY = 1 - mvarB / 255
    mvarK = mvarC
    If mvarM < mvarK Then mvarK = mvarM
    If mvarY < mvarK Then mvarK = mvarY
    If mvarK = 1 Then
        mvarC = 0: mvarM = 0: mvarY = 0
    Else
        mvarC = (mvarC - mvarK) / (1 - mvarK)
        mvarM = (mvarM - mvarK) / (1 - mvarK)
        mvarY = (mvarY - mvarK) / (1 - mvarK)
    End If
End Sub

