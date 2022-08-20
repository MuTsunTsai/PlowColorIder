VERSION 5.00
Begin VB.UserControl PlowRGB 
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   238
   ScaleMode       =   3  '像素
   ScaleWidth      =   320
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   60
      ScaleHeight     =   21
      ScaleMode       =   3  '像素
      ScaleWidth      =   293
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4455
   End
   Begin VB.Image I 
      Height          =   90
      Left            =   240
      Picture         =   "PlowRGB.ctx":0000
      Top             =   480
      Width           =   165
   End
End
Attribute VB_Name = "PlowRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private MX As Single

'Default Property Values:
Const m_def_R = 0
Const m_def_G = 0
Const m_def_B = 0
Const m_def_Value = 0

'Property Variables:
Dim m_R As Integer
Dim m_G As Integer
Dim m_B As Integer
Dim m_Value As Integer

'Event Declarations:
Event Change(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=I,I,-1,MouseUp
Attribute Change.VB_Description = "發生於使用者在物件具有駐點 (focus) 時，放開滑鼠鍵。"

Public Sub Setvalue(ByVal new_R As Integer, ByVal new_G As Integer, ByVal new_B As Integer)
    m_R = new_R: m_G = new_G: m_B = new_B
    Draw
End Sub

Private Sub Draw()
    Static V(1) As TRIVERTEX, G As GRADIENT_RECT
    
    With V(0)
        .X = 0
        .Y = 0
        .Red = IIf(m_R = 256, 0, Color16(m_R))
        .Green = IIf(m_G = 256, 0, Color16(m_G))
        .Blue = IIf(m_B = 256, 0, Color16(m_B))
    End With
    With V(1)
        .X = P.ScaleWidth
        .Y = P.ScaleHeight
        .Red = IIf(m_R = 256, &HFF00, Color16(m_R))
        .Green = IIf(m_G = 256, &HFF00, Color16(m_G))
        .Blue = IIf(m_B = 256, &HFF00, Color16(m_B))
    End With
    
    G.UpperLeft = 0
    G.LowerRight = 1
       
    GradientFill P.hdc, V(0), 2, G, 1, GRADIENT_FILL_RECT_H
    P.Refresh
End Sub

Private Sub I_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MX = X / Screen.TwipsPerPixelX
End Sub

Private Sub I_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        I.Left = ValueInRange(I.Left + X / Screen.TwipsPerPixelX - MX, 0, P.ScaleWidth)
        m_Value = I.Left / P.ScaleWidth * 255
        RaiseEvent Change(Button, Shift, X, Y)
    End If
End Sub

Private Sub P_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    P_MouseMove Button, Shift, X, Y
End Sub

Private Sub P_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        I.Left = ValueInRange(X, 0, P.ScaleWidth)
        m_Value = I.Left / P.ScaleWidth * 255
        RaiseEvent Change(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_Resize()
    P.Width = UserControl.ScaleWidth - 6
    P.Height = UserControl.ScaleHeight - 6
    I.Left = 0
    I.Top = P.Height
End Sub

'警告! 切勿移除或修改以下的註解行!
'MemberInfo=7,0,0,0
Public Property Get R() As Integer
    R = m_R
End Property

Public Property Let R(ByVal new_R As Integer)
    m_R = new_R
    PropertyChanged "R"
    Draw
End Property

'警告! 切勿移除或修改以下的註解行!
'MemberInfo=7,0,0,0
Public Property Get G() As Integer
    G = m_G
End Property

Public Property Let G(ByVal new_G As Integer)
    m_G = new_G
    PropertyChanged "G"
    Draw
End Property

'警告! 切勿移除或修改以下的註解行!
'MemberInfo=7,0,0,0
Public Property Get B() As Integer
    B = m_B
End Property

Public Property Let B(ByVal new_B As Integer)
    m_B = new_B
    PropertyChanged "B"
    Draw
End Property

'警告! 切勿移除或修改以下的註解行!
'MemberInfo=7,0,0,0
Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    I.Left = (m_Value / 255) * P.ScaleWidth
    PropertyChanged "Value"
End Property

'初始化使用者控制項的屬性
Private Sub UserControl_InitProperties()
    m_R = m_def_R
    m_G = m_def_G
    m_B = m_def_B
    m_Value = m_def_Value
End Sub

'由儲存區載入屬性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_R = PropBag.ReadProperty("R", m_def_R)
    m_G = PropBag.ReadProperty("G", m_def_G)
    m_B = PropBag.ReadProperty("B", m_def_B)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
End Sub

Private Sub UserControl_Show()
    P.ToolTipText = Extender.ToolTipText
    I.ToolTipText = Extender.ToolTipText
End Sub

'將屬性值寫回儲存區
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("R", m_R, m_def_R)
    Call PropBag.WriteProperty("G", m_G, m_def_G)
    Call PropBag.WriteProperty("B", m_B, m_def_B)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub

'警告! 切勿移除或修改以下的註解行!
'MappingInfo=P,P,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "傳回物件視窗的物件代碼 (由 Microsoft Windows 所提供)。"
    hWnd = P.hWnd
End Property

