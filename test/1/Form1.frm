VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  '螢幕中央
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      Height          =   1815
      Left            =   360
      ScaleHeight     =   117
      ScaleMode       =   3  '像素
      ScaleWidth      =   117
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private C As Long
Private DIB As New DIBSection

Private Sub Form_Load()
    Static K As Long, T As String
    Static I As Integer
    
    SetStretchBltMode P1.hDC, 3
       
    'T = Now
    'For K = 0 To 10000
    '    Draw5
    'Next K
    
    'Debug.Print DateDiff("s", T, Now)
End Sub

' 方法零：沒做事
' 需時為 0.00037, 0.00043 秒

Private Sub Draw0()
    Static C As New ColorHSV
    Static J As Double, H As Integer, I As Integer
    
    C.S = 100: C.H = 150
    J = P1.ScaleHeight / 101: H = P1.ScaleWidth
    
    For I = 0 To 100
        C.I = 100 - I
    Next I
End Sub

' 方法一：使用 Line
' 需時為 0.0028, 0.0029 秒

Private Sub Draw1()
    Static C As New ColorHSV
    Static J As Double, H As Integer, I As Integer
    
    C.S = 100: C.H = 150
    J = P1.ScaleHeight / 101: H = P1.ScaleWidth
    
    For I = 0 To 100
        C.I = 100 - I
        P1.Line (0, I * J)-(H, (I + 1) * J - 1), C.Color, BF
    Next I
End Sub

' 方法二：使用 FillRect
' 需時為 0.0011, 0.0011 秒

Private Sub Draw2()
    Static C As New ColorHSV
    Static J As Double, H As Integer, I As Integer
    Static B As Long
    
    C.S = 100: C.H = 150
    J = P1.ScaleHeight / 101: H = P1.ScaleWidth
      
    For I = 0 To 100
        C.I = 100 - I
        B = CreateSolidBrush(C.Color)
        FillRect P1.hDC, CreateRect(0, I * J, H, (I + 1) * J), B
        DeleteObject B
    Next I
End Sub

' 方法三：使用 GradientFill
' 需時為 0.00057, 0.0011 秒

Public Function Color16(Clr As Byte) As Integer
    Dim UInt As Long
    UInt = Clr * &H100&
    If UInt < &H7FFF Then
        Color16 = CInt(UInt)
    Else
        Color16 = CInt(UInt - &H10000)
    End If
End Function

Private Sub SetColor(ByRef V As TRIVERTEX, ByVal C As Long)
    V.Red = Color16(C And &HFF)
    V.Green = Color16((C \ &H100) And &HFF)
    V.Blue = Color16((C \ &H10000) And &HFF)
End Sub

Private Sub Draw3()
    Static C As New ColorHSV
    Static G As GRADIENT_RECT
    Static V(1) As TRIVERTEX
    
    C.S = 100: C.H = 150: C.I = 100
    
    V(0).X = 0: V(0).Y = 0: SetColor V(0), C.Color
    V(1).X = P1.ScaleWidth: V(1).Y = P1.ScaleHeight
    SetColor V(1), 0
    
    G.UpperLeft = 0: G.LowerRight = 1
      
    GradientFill P1.hDC, V(0), 2, G, 1, GRADIENT_FILL_RECT_V
End Sub

Private Sub Draw4()
    Static C As New ColorHSL
    Static G As GRADIENT_RECT
    Static V(1) As TRIVERTEX
    
    C.S = 100: C.H = 150
    
    G.UpperLeft = 0: G.LowerRight = 1
    
    V(0).X = 0: V(0).Y = 0
    C.I = 100: SetColor V(0), C.Color
    V(1).X = P1.ScaleWidth: V(1).Y = P1.ScaleHeight / 2
    C.I = 50: SetColor V(1), C.Color
    GradientFill P1.hDC, V(0), 2, G, 1, GRADIENT_FILL_RECT_V
    
    V(0).X = 0: V(0).Y = P1.ScaleHeight / 2
    C.I = 50: SetColor V(0), C.Color
    V(1).X = P1.ScaleWidth: V(1).Y = P1.ScaleHeight
    SetColor V(1), 0
    GradientFill P1.hDC, V(0), 2, G, 1, GRADIENT_FILL_RECT_V
    
End Sub

' 方法四：使用 DIBSection
' 需時為 0.0028 秒

Private Sub Draw5()
    Static I As Integer, J As Double, H As Integer
    Static C As New ColorHSV
    
    H = 100
    
    DIB.CreateDIB Me.hDC, 1, H + 1
    
    C.S = 100: C.H = 150
    C.I = 100
    
    For J = 0 To H
        C.I = J / H * 100
        DIB.SetPoint C.Color, 0, J
    Next J
    
    StretchBlt P1.hDC, 0, 0, P1.ScaleWidth, P1.ScaleHeight, DIB.hDC, 0, 0, 1, H + 1, vbSrcCopy
    
End Sub

Private Sub P1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
             
    If Button = 1 Then
        PaintHS Y
        Debug.Print Time
    End If

End Sub

Private Sub PaintSV(ByVal Y As Integer)
    Static I As Integer, J As Double, H As Integer
    
    Dim fS As Double, fI As Double, fH As Double
    Dim nI As Integer
    Dim fF As Double, fP As Double, fQ As Double, fT As Double
    Dim SR As Double, SG As Double, SB As Double

    H = P1.ScaleHeight - 1
    fI = 1
    
    fH = Y / H * 359 / 60
    If fH = 6 Then fH = 0
    nI = Int(fH)
    fF = fH - nI
    
    For I = 0 To 4 * (P1.ScaleWidth - 1) Step 4
        
        fS = I / 4 / (P1.ScaleHeight - 1)

        If fS = 0 Then
            SR = 1
            SG = 1
            SB = 1
        Else
            fP = 1 - fS
            fQ = 1 - (fS * fF)
            fT = 1 - (fS * (1 - fF))
            
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
        
        'B(I, H) = 255 * SB
        'B(I + 1, H) = 255 * SG
        'B(I + 2, H) = 255 * SR
    Next I
    
    For J = 0 To H - 1
        For I = 0 To 4 * (P1.ScaleWidth - 1) Step 4
            'B(I, J) = B(I, H) * J / H
            'B(I + 1, J) = B(I + 1, H) * J / H
            'B(I + 2, J) = B(I + 2, H) * J / H
        Next I
    Next J
    
    BitBlt P1.hDC, 0, 0, P1.ScaleWidth, P1.ScaleHeight, DC, 0, 0, vbSrcCopy
    P1.Refresh
        
End Sub

Private Sub PaintHS(ByVal Y As Integer)
    Static I As Integer, J As Double
    Static H As Integer, W As Integer
    
    H = 50
    W = 180
      
    DIB.CreateDIB Me.hDC, W + 1, H + 1
    
    Dim fS As Double, fI As Double, fH As Double
    Dim nI As Integer
    Dim fF As Double, fP As Double, fQ As Double, fT As Double
    Dim SR As Double, SG As Double, SB As Double
    
    fI = Y / (P1.ScaleHeight - 1)
    If fI < 0 Then fI = 0
    If fI > 1 Then fI = 1
    
    For J = 0 To H
        fS = J / H
    
        For I = 0 To W
            
            fH = I / W * 359 / 60
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
                   
            DIB.SetPointRGB 255 * SR, 255 * SG, 255 * SB, I, J
        Next I
        
    Next J
    StretchBlt P1.hDC, 0, 0, P1.ScaleWidth, P1.ScaleHeight, DIB.hDC, 0, 0, W + 1, H + 1, vbSrcCopy
    P1.Refresh
        
End Sub

