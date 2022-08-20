VERSION 5.00
Begin VB.UserControl PlowColorB 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MousePointer    =   10  '往上指
   ScaleHeight     =   240
   ScaleMode       =   3  '像素
   ScaleWidth      =   320
   Begin VB.Image I 
      Height          =   165
      Index           =   2
      Left            =   600
      Picture         =   "PlowColorB.ctx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image I 
      Height          =   165
      Index           =   1
      Left            =   360
      Picture         =   "PlowColorB.ctx":0357
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image I 
      Height          =   165
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   165
   End
End
Attribute VB_Name = "PlowColorB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Const m_def_Style = "HI"
Private Const m_def_Color = 0

Private m_Style As String
Private m_Color As Long

Private Sub I_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X / Screen.TwipsPerPixelX + I(Index).Left, Y / Screen.TwipsPerPixelY + I(Index).Top
End Sub

Private Sub I_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X / Screen.TwipsPerPixelX + I(Index).Left, Y / Screen.TwipsPerPixelY + I(Index).Top
End Sub

Private Sub I_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X / Screen.TwipsPerPixelX + I(Index).Left, Y / Screen.TwipsPerPixelY + I(Index).Top
End Sub

Private Sub UpdateColor(ByVal X As Single, ByVal Y As Single)
    Static C As Object, rX As Double, rY As Double
    
    rX = X / (UserControl.ScaleWidth - 1)
    rY = 1 - Y / (UserControl.ScaleHeight - 1)
    
    If m_Style = "HI" Then
        Set C = New ColorHSL
        C.H = rX * 359: C.S = 100: C.I = rY * 100
    Else
        If HSVMode Then
            Set C = New ColorHSV
            C.I = 100
        Else
            Set C = New ColorHSL
            C.I = 50
        End If
        C.H = rX * 359: C.S = rY * 100
    End If
    m_Color = C.Color
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static C As Object
    If Button = 1 Then
        UpdateColor X, Y
        Set C = New ColorHSL
        C.Color = m_Color
        If (m_Style = "HI" And C.I >= 50) Or (m_Style <> "HI" And HSVMode) Then _
            I(0).Picture = I(2).Picture Else I(0).Picture = I(1).Picture
        I(0).Left = X - 5: I(0).Top = Y - 5
        LockCursor UserControl.hWnd, 2
        If m_Style = "HI" Then I(0).Visible = True
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static C As Object
    If Button = 1 Then
        UpdateColor X, Y
        Set C = New ColorHSL
        C.Color = m_Color
        If (m_Style = "HI" And C.I >= 50) Or (m_Style <> "HI" And HSVMode) Then _
            I(0).Picture = I(2).Picture Else I(0).Picture = I(1).Picture
        I(0).Left = X - 5: I(0).Top = Y - 5
        LockCursor UserControl.hWnd, 2
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then UnLockCursor: If m_Style = "HI" Then I(0).Visible = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Function Point(X As Single, Y As Single) As Long
    Point = UserControl.Point(X, Y)
End Function

Public Sub Draw()
    Static C As Object
    Static cImage As New c32bppDIB
    Static W, H, I As Long, j As Long, K As Long
    DoEvents
    If m_Style = "HI" Or m_Style = "" Then
        '由於本選項圖片不會再變動，為了提昇執行效能，以資源圖片的方式取代程式繪圖
        cImage.LoadPicture_Resource 101, "JPEG"
        cImage.Render UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        '底下是原本用以繪圖的程式碼
        'Set C = New ColorHSL: C.S = 100
        'W = UserControl.ScaleWidth - 1: H = UserControl.ScaleHeight - 1
        'For I = 0 To W
        '    C.H = I / W * 359
        '    For j = 0 To H
        '        C.I = (H - j) / H * 100
        '        SetPixel UserControl.hDC, I, j, C.Color
        '        'UserControl.PSet (I, J), C.Color
        '    Next j
        'Next I
        UserControl.Refresh
    ElseIf m_Style = "HS" Then
        '由於本選項圖片不會再變動，為了提昇執行效能，以資源圖片的方式取代程式繪圖
        If HSVMode Then
            cImage.LoadPicture_Resource 102, "JPEG"
            cImage.Render UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        Else
            cImage.LoadPicture_Resource 103, "JPEG"
            cImage.Render UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        End If
        '底下是原本用以繪圖的程式碼
        'If HSVMode Then Set C = New ColorHSV: C.I = 100 Else Set C = New ColorHSL: C.I = 50
        'W = UserControl.ScaleWidth - 1: H = UserControl.ScaleHeight - 1
        'For I = 0 To W
        '    C.H = I / W * 359
        '    For J = 0 To H
        '        C.S = (H - J) / H * 100
        '        SetPixel UserControl.hdc, I / 1, J / 1, C.Color
        '        'UserControl.PSet (I, J), C.Color
        '    Next J
        'Next I
        UserControl.Refresh
    '原本有計劃設計逆向色盤，但因執行速度太差而放棄
    'ElseIf m_Style = "SI" Then
        'UserControl.Picture = LoadResPicture(104, vbResBitmap)
        'Picture2Array UserControl.Picture, F()
        'If HSVMode Then Set C = New ColorHSV Else Set C = New ColorHSL
        'C.H = m_H
        'For I = 0 To 116
        '    C.S = I / 117 * 100
        '    For J = 0 To 116
        '        C.I = (117 - J) / 117 * 100
        '        K = C.Color
        '        F((117 * I + J) * 3 + 172) = K And &HFF
        '        F((117 * I + J) * 3 + 173) = (K \ &H100) And &HFF
        '        F((117 * I + J) * 3 + 174) = (K \ &H10000) And &HFF
        '   Next J
        'Next I
        'UserControl.PaintPicture Array2Picture(F), 0, 0
        'If HSVMode Then Set C = New ColorHSV Else Set C = New ColorHSL
        'C.H = m_H
        'W = UserControl.ScaleWidth: H = UserControl.ScaleHeight
        'For I = 0 To W
        '    C.S = I / W * 100
        '    For J = 0 To H
        '        C.I = (H - J) / H * 100
        '        UserControl.PSet (I, J), C.Color
        '    Next J
        'Next I
        'Picture2Array UserControl.Image, F
        'W = ""
        'For I = 0 To UBound(F)
        '    W = W & Hex(F(I))
        'Next I
        'Stop
        'MsgBox UBound(F)
        'MsgBox W
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
'    m_H = PropBag.ReadProperty("H", m_def_H)
    m_Color = PropBag.ReadProperty("Color", m_def_Color)
End Sub

Private Sub UserControl_Show()
    I(0).ToolTipText = Extender.ToolTipText
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("Color", m_Color, m_def_Color)
End Sub

Public Property Get Style() As String
Attribute Style.VB_Description = "設定色盤的型態。"
    If Ambient.UserMode Then Err.Raise 393
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As String)
    If Ambient.UserMode Then Err.Raise 382
    m_Style = New_Style
    PropertyChanged "Style"
End Property

Private Sub UserControl_InitProperties()
    m_Style = m_def_Style
    m_Color = m_def_Color
End Sub

Public Property Get Color() As Long
    Color = m_Color
End Property

Public Property Let Color(ByVal New_Color As Long)
    Static C As Object
    If HSVMode Then Set C = New ColorHSV Else Set C = New ColorHSL
    
    m_Color = New_Color
    C.Color = m_Color
    If (m_Style = "HI" And C.I >= 50) Or (m_Style <> "HI" And HSVMode) Then _
        I(0).Picture = I(2).Picture Else I(0).Picture = I(1).Picture
    I(0).Left = C.H / 360 * UserControl.ScaleWidth - 5
    I(0).Top = (100 - C.S) / 101 * UserControl.ScaleHeight - 5
    PropertyChanged "Color"
End Property

Public Sub PointerMove(ByVal X As Integer, ByVal Y As Integer)
    I(0).Left = ValueInRange(I(0).Left + X, -5, UserControl.ScaleWidth - 6)
    I(0).Top = ValueInRange(I(0).Top + Y, -5, UserControl.ScaleHeight - 6)
    UpdateColor I(0).Left + 5, I(0).Top + 5
End Sub
