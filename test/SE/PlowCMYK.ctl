VERSION 5.00
Begin VB.UserControl PlowCMYK 
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   ScaleHeight     =   2715
   ScaleWidth      =   4515
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   45
      ScaleHeight     =   315
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Image I 
      Height          =   90
      Left            =   315
      Picture         =   "PlowCMYK.ctx":0000
      Top             =   720
      Width           =   165
   End
End
Attribute VB_Name = "PlowCMYK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_C = 0
Const m_def_M = 0
Const m_def_Y = 0
Const m_def_K = 0
Const m_def_Value = 0
'Property Variables:
Dim m_C As Integer
Dim m_M As Integer
Dim m_Y As Integer
Dim m_K As Integer
Dim m_Value As Integer

'Event Declarations:
Event Change(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=I,I,-1,MouseUp
Attribute Change.VB_Description = "發生於使用者在物件具有駐點 (focus) 時，放開滑鼠鍵。"

Public Sub Setvalue(ByVal new_C As Integer, ByVal new_M As Integer, ByVal new_Y As Integer, ByVal new_K As Integer)
    m_C = new_C: m_M = new_M: m_Y = new_Y: m_K = new_K
    Draw
End Sub

Private Sub Draw()
    Static H As Integer, W As Integer, I As Integer
    Static CC As New ColorCMYK
    H = P.ScaleHeight: W = P.ScaleWidth
    If m_C = 101 Then
        CC.M = m_M: CC.Y = m_Y: CC.K = m_K
        For I = 0 To W Step 15
            CC.C = I / W * 100
            P.Line (I, 0)-(I, H), CC.Color
        Next I
    ElseIf m_M = 101 Then
        CC.C = m_C: CC.Y = m_Y: CC.K = m_K
        For I = 0 To W Step 15
            CC.M = I / W * 100
            P.Line (I, 0)-(I, H), CC.Color
        Next I
    ElseIf m_Y = 101 Then
        CC.C = m_C: CC.M = m_M: CC.K = m_K
        For I = 0 To W Step 15
            CC.Y = I / W * 100
            P.Line (I, 0)-(I, H), CC.Color
        Next I
    ElseIf m_K = 101 Then
        CC.C = m_C: CC.M = m_M: CC.Y = m_Y
        For I = 0 To W Step 15
            CC.K = I / W * 100
            P.Line (I, 0)-(I, H), CC.Color
        Next I
    End If
End Sub

Private Sub I_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MX = X
End Sub

Private Sub I_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        I.Move I.Left + X - MX
        If I.Left < 0 Then I.Left = 0
        If I.Left > P.ScaleWidth Then I.Left = P.ScaleWidth
        
        m_Value = CInt((I.Left / P.ScaleWidth * 100) / 10) * 10
        I.Left = (m_Value / 100) * P.ScaleWidth
        
        RaiseEvent Change(Button, Shift, X, Y)
    End If
End Sub

Private Sub P_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    P_MouseMove Button, Shift, X, Y
End Sub

Private Sub P_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        I.Move X
        If I.Left < 0 Then I.Left = 0
        If I.Left > P.ScaleWidth Then I.Left = P.ScaleWidth
        m_Value = I.Left / P.ScaleWidth * 100
        
        m_Value = CInt((I.Left / P.ScaleWidth * 100) / 10) * 10
        I.Left = (m_Value / 100) * P.ScaleWidth
        
        RaiseEvent Change(Button, Shift, X, Y)
    End If
End Sub


'警告! 切勿移除或修改以下的註解行!
'MemberInfo=7,0,0,0
Public Property Get C() As Integer
    C = m_C
End Property

Public Property Let C(ByVal new_C As Integer)
    m_C = new_C
    PropertyChanged "C"
    Draw
End Property

'警告! 切勿移除或修改以下的註解行!
'MemberInfo=7,0,0,0
Public Property Get M() As Integer
    M = m_M
End Property

Public Property Let M(ByVal new_M As Integer)
    m_M = new_M
    PropertyChanged "M"
    Draw
End Property

'警告! 切勿移除或修改以下的註解行!
'MemberInfo=7,0,0,0
Public Property Get Y() As Integer
    Y = m_Y
End Property

Public Property Let Y(ByVal new_Y As Integer)
    m_Y = new_Y
    PropertyChanged "Y"
    Draw
End Property

'警告! 切勿移除或修改以下的註解行!
'MemberInfo=7,0,0,0
Public Property Get K() As Integer
    K = m_K
End Property

Public Property Let K(ByVal new_K As Integer)
    m_K = new_K
    PropertyChanged "K"
    Draw
End Property

'警告! 切勿移除或修改以下的註解行!
'MemberInfo=7,0,0,0
Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    I.Left = (m_Value / 100) * P.ScaleWidth
    PropertyChanged "Value"
End Property

'初始化使用者控制項的屬性
Private Sub UserControl_InitProperties()
    m_C = m_def_C
    m_M = m_def_M
    m_Y = m_def_Y
    m_K = m_def_K
    m_Value = m_def_Value
End Sub

'由儲存區載入屬性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_C = PropBag.ReadProperty("C", m_def_C)
    m_M = PropBag.ReadProperty("M", m_def_M)
    m_Y = PropBag.ReadProperty("Y", m_def_Y)
    m_K = PropBag.ReadProperty("K", m_def_K)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
End Sub

Private Sub UserControl_Resize()
    P.Width = UserControl.ScaleWidth - 90
    P.Height = UserControl.ScaleHeight - 90
    I.Left = 0
    I.Top = P.Height
End Sub

Private Sub UserControl_Show()
    P.ToolTipText = Extender.ToolTipText
    I.ToolTipText = Extender.ToolTipText
End Sub

'將屬性值寫回儲存區
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("C", m_C, m_def_C)
    Call PropBag.WriteProperty("M", m_M, m_def_M)
    Call PropBag.WriteProperty("Y", m_Y, m_def_Y)
    Call PropBag.WriteProperty("K", m_K, m_def_K)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub

Private Sub I_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent Change(Button, Shift, X, Y)
End Sub

