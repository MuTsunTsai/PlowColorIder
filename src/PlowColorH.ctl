VERSION 5.00
Begin VB.UserControl PlowColorH 
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
   Begin VB.Image Img 
      Height          =   165
      Index           =   0
      Left            =   60
      Top             =   120
      Width           =   165
   End
   Begin VB.Image Img 
      Height          =   165
      Index           =   1
      Left            =   360
      Picture         =   "PlowColorH.ctx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Img 
      Height          =   165
      Index           =   2
      Left            =   600
      Picture         =   "PlowColorH.ctx":0357
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "PlowColorH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Const m_def_H = 0
Const m_def_S = 0
Const m_def_I = 0
Const m_def_Style = "I"

Private m_H As Integer
Private m_S As Integer
Private m_Style As String
Private m_I As Integer


Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property



Private Sub SwitchImage(ByVal Color As Long)
    Static C As New ColorHSL
    C.Color = Color
    If C.I > 50 Then Img(0).Picture = Img(2).Picture Else Img(0).Picture = Img(1).Picture
End Sub




Private Sub Img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X / Screen.TwipsPerPixelX + Img(Index).Left, Y / Screen.TwipsPerPixelY + Img(Index).Top
End Sub

Private Sub Img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X / Screen.TwipsPerPixelX + Img(Index).Left, Y / Screen.TwipsPerPixelY + Img(Index).Top
End Sub

Private Sub Img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X / Screen.TwipsPerPixelX + Img(Index).Left, Y / Screen.TwipsPerPixelY + Img(Index).Top
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SwitchImage UserControl.Point(X, Y)
        Img(0).Top = Y - 5
        m_I = 100 - CInt(100 * (Img(0).Top + 5) / UserControl.ScaleHeight)
        LockCursor UserControl.hWnd, 2
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static C As Object
    If Button = 1 Then
        SwitchImage UserControl.Point(X, Y)
        Img(0).Top = Y - 5
        m_I = 100 - CInt(100 * (Img(0).Top + 5) / UserControl.ScaleHeight)
        LockCursor UserControl.hWnd, 2
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then UnLockCursor
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub Draw()
    Static C As Object
    Static G As GRADIENT_RECT
    Static V(1) As TRIVERTEX
    
    If m_Style = "I" Or m_Style = "" Then
    
        G.UpperLeft = 0: G.LowerRight = 1
        If HSVMode Then
            Set C = New ColorHSV
            C.S = m_S: C.H = m_H: C.I = 100
            V(0).X = 0: V(0).Y = 0: SetTriVertexColor V(0), C.Color
        Else
            Set C = New ColorHSL
            C.S = m_S: C.H = m_H
            V(0).X = 0: V(0).Y = 0
            C.I = 100: SetTriVertexColor V(0), C.Color
            V(1).X = UserControl.ScaleWidth: V(1).Y = UserControl.ScaleHeight / 2 - 1
            C.I = 50: SetTriVertexColor V(1), C.Color
            GradientFill UserControl.hDC, V(0), 2, G, 1, GRADIENT_FILL_RECT_V
            V(0).X = 0: V(0).Y = UserControl.ScaleHeight / 2
            SetTriVertexColor V(0), C.Color
        End If
        V(1).X = UserControl.ScaleWidth: V(1).Y = UserControl.ScaleHeight - 1
        SetTriVertexColor V(1), 0
        GradientFill UserControl.hDC, V(0), 2, G, 1, GRADIENT_FILL_RECT_V
        UserControl.Line (0, UserControl.ScaleHeight - 1)- _
            (UserControl.ScaleWidth, UserControl.ScaleHeight - 1), vbBlack
        
        
    '原本有計劃設計逆向色盤，但因執行速度太差而放棄
    'ElseIf m_Style = "H" Then
        'Set C = New ColorHSV
        'C.S = 100: C.I = 100
        'J = UserControl.ScaleHeight / 360: H = UserControl.ScaleWidth
        'For I = 0 To 359
        '    C.H = 359 - I
        '    UserControl.Line (0, I * J)-(H, (I + 1) * J - 1), C.Color, BF
        'Next I
        
    End If
    SwitchImage Now
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_H = PropBag.ReadProperty("H", m_def_H)
    m_S = PropBag.ReadProperty("S", m_def_S)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
End Sub

Private Sub UserControl_Show()
    Img(0).ToolTipText = Extender.ToolTipText
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("H", m_H, m_def_H)
    Call PropBag.WriteProperty("S", m_S, m_def_S)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
End Sub

Public Property Get H() As Integer
    H = m_H
End Property

Public Property Let H(ByVal New_H As Integer)
    Static C As Object
    m_H = New_H
    PropertyChanged "H"
End Property

Public Property Get S() As Integer
    S = m_S
End Property

Public Property Let S(ByVal New_S As Integer)
    Static C As Object
    m_S = New_S
    PropertyChanged "S"
End Property

Private Sub UserControl_InitProperties()
    m_H = m_def_H
    m_S = m_def_S
    m_I = m_def_I
    m_Style = m_def_Style
End Sub

Public Property Get Style() As String
    If Ambient.UserMode Then Err.Raise 393
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As String)
    If Ambient.UserMode Then Err.Raise 382
    m_Style = New_Style
    PropertyChanged "Style"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Color() As Long
    Color = Now
End Property

Public Property Let Color(ByVal New_Color As Long)
    Static C As Object
    If HSVMode Then Set C = New ColorHSV Else Set C = New ColorHSL
    
    C.Color = New_Color
    m_H = C.H: m_S = C.S: m_I = C.I
    Img(0).Top = CInt((100 - m_I) / 101 * UserControl.ScaleHeight) - 5
    SwitchImage New_Color
    
    PropertyChanged "Color"
End Property

Private Function Now() As Long
    Now = UserControl.Point(Img(0).Left + 5, Img(0).Top + 5)
End Function

Public Property Get I() As Integer
    I = m_I
End Property

Public Property Let I(ByVal New_I As Integer)
    Static C As Object
    m_I = ValueInRange(New_I, 0, 100)
    Img(0).Top = CInt((100 - m_I) / 101 * UserControl.ScaleHeight) - 5
    SwitchImage Now
End Property


