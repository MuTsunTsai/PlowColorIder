VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorHSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private mlColor As Long, R As Double, G As Double, B As Double
Private SH As Double, SS As Double, SI As Double, OI As Double

Public Property Let Color(lColor As Long)
    mlColor = lColor
    R = mlColor And &HFF
    G = (mlColor \ &H100) And &HFF
    B = (mlColor \ &H10000) And &HFF
    ToSHsv
End Property

Public Property Get Color() As Long
    Color = RGB(CInt(R), CInt(G), CInt(B))
End Property

Public Property Let H(L As Double)
    SH = L
    ToRgb
End Property

Public Property Let S(L As Double)
    SS = L
    ToRgb
End Property

Public Property Let I(L As Double)
    SI = ValueInRange(L, 0, 100)
    ChangeI
    OI = SI
End Property

Public Property Get H() As Double
    H = SH
End Property

Public Property Get S() As Double
    S = SS
End Property

Public Property Get I() As Double
    I = SI
End Property

Private Sub ToSHsv()
    Static SR As Double, SG As Double, SB As Double
    
    Dim fMx As Double
    Dim fMn As Double
    Dim fVa As Double
    Dim fSa As Double
    Dim fRc As Double
    Dim fGc As Double
    Dim fBc As Double
    
    SR = R / 255: SG = G / 255: SB = B / 255
    
    fMx = SR
    If SG > fMx Then fMx = SG
    If SB > fMx Then fMx = SB
    fMn = SR
    If SG < fMn Then fMn = SG
    If SB < fMn Then fMn = SB
    fVa = fMx
    
    If fMx Then
        fSa = (fMx - fMn) / fMx
    Else
        fSa = 0
    End If
    
    If fSa = 0 Then
        SH = 0
    Else
        fRc = (fMx - SR) / (fMx - fMn)
        fGc = (fMx - SG) / (fMx - fMn)
        fBc = (fMx - SB) / (fMx - fMn)
        Select Case fMx
        Case SR
            SH = fBc - fGc
        Case SG
            SH = 2 + fRc - fBc
        Case SB
            SH = 4 + fGc - fRc
        End Select
        SH = SH * 60
        If SH < 0 Then SH = SH + 360
    End If
    SS = fSa * 100
    SI = fVa * 100
    
    OI = SI

End Sub

Private Sub ToRgb()
    Dim fS As Double
    Dim fI As Double
    Dim fH As Double
    Dim nI As Integer
    Dim fF As Double
    Dim fP As Double
    Dim fQ As Double
    Dim fT As Double
    Dim SR As Double
    Dim SG As Double
    Dim SB As Double
    
    fS = SS / 100
    fI = SI / 100
    
    If SS = 0 Then
        SR = fI
        SG = fI
        SB = fI
    Else
        fH = SH / 60
        If fH = 6 Then fH = 0
        nI = Int(fH)
        fF = fH - nI
        
        fP = fI * (1 - fS)
        fQ = fI * (1 - (fS * fF))
        fT = fI * (1 - (fS * (1 - fF)))
        
        Select Case nI
            Case 0
                SR = fI
                SG = fT
                SB = fP
            Case 1
                SR = fQ
                SG = fI
                SB = fP
            Case 2
                SR = fP
                SG = fI
                SB = fT
            Case 3
                SR = fP
                SG = fQ
                SB = fI
            Case 4
                SR = fT
                SG = fP
                SB = fI
            Case 5
                SR = fI
                SG = fP
                SB = fQ
        End Select
    End If
    
    R = 255 * SR
    G = 255 * SG
    B = 255 * SB
        
End Sub

Private Sub ChangeI()
    If OI <> 0 Then
        R = R / OI * SI
        G = G / OI * SI
        B = B / OI * SI
    Else
        ToRgb
    End If
End Sub

