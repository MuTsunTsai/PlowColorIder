VERSION 5.00
Begin VB.UserControl PlowCMYK 
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   ScaleHeight     =   181
   ScaleMode       =   3  '����
   ScaleWidth      =   301
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   45
      ScaleHeight     =   21
      ScaleMode       =   3  '����
      ScaleWidth      =   269
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

Option Explicit

Private MX As Single

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
Attribute Change.VB_Description = "�o�ͩ�ϥΪ̦b����㦳�n�I (focus) �ɡA��}�ƹ���C"

Public Sub Setvalue(ByVal new_C As Integer, ByVal new_M As Integer, ByVal new_Y As Integer, ByVal new_K As Integer)
    m_C = new_C: m_M = new_M: m_Y = new_Y: m_K = new_K
    Draw
End Sub

Private Sub Draw()
    Static V(1) As TRIVERTEX, G As GRADIENT_RECT, CC As New ColorCMYK
    
    CC.C = IIf(m_C = 101, 0, m_C)
    CC.m = IIf(m_M = 101, 0, m_M)
    CC.Y = IIf(m_Y = 101, 0, m_Y)
    CC.K = IIf(m_K = 101, 0, m_K)
    V(0).X = 0: V(0).Y = 0
    SetTriVertexColor V(0), CC.Color
    
    CC.C = IIf(m_C = 101, 100, m_C)
    CC.m = IIf(m_M = 101, 100, m_M)
    CC.Y = IIf(m_Y = 101, 100, m_Y)
    CC.K = IIf(m_K = 101, 100, m_K)
    V(1).X = P.ScaleWidth: V(1).Y = P.ScaleHeight
    SetTriVertexColor V(1), CC.Color
    
    G.UpperLeft = 0: G.LowerRight = 1
    GradientFill P.hDC, V(0), 2, G, 1, GRADIENT_FILL_RECT_H
    P.Refresh

End Sub

Private Sub I_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MX = X / Screen.TwipsPerPixelX
End Sub

Private Sub I_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        I.Left = ValueInRange(I.Left + X / Screen.TwipsPerPixelX - MX, 0, P.ScaleWidth)
        m_Value = I.Left / P.ScaleWidth * 100
        RaiseEvent Change(Button, Shift, X, Y)
    End If
End Sub

Private Sub P_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    P_MouseMove Button, Shift, X, Y
End Sub

Private Sub P_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        I.Left = ValueInRange(X, 0, P.ScaleWidth)
        m_Value = I.Left / P.ScaleWidth * 100
        RaiseEvent Change(Button, Shift, X, Y)
    End If
End Sub


'ĵ�i! ���Ų����έק�H�U�����Ѧ�!
'MemberInfo=7,0,0,0
Public Property Get C() As Integer
    C = m_C
End Property

Public Property Let C(ByVal new_C As Integer)
    m_C = new_C
    PropertyChanged "C"
    Draw
End Property

'ĵ�i! ���Ų����έק�H�U�����Ѧ�!
'MemberInfo=7,0,0,0
Public Property Get m() As Integer
    m = m_M
End Property

Public Property Let m(ByVal new_M As Integer)
    m_M = new_M
    PropertyChanged "M"
    Draw
End Property

'ĵ�i! ���Ų����έק�H�U�����Ѧ�!
'MemberInfo=7,0,0,0
Public Property Get Y() As Integer
    Y = m_Y
End Property

Public Property Let Y(ByVal new_Y As Integer)
    m_Y = new_Y
    PropertyChanged "Y"
    Draw
End Property

'ĵ�i! ���Ų����έק�H�U�����Ѧ�!
'MemberInfo=7,0,0,0
Public Property Get K() As Integer
    K = m_K
End Property

Public Property Let K(ByVal new_K As Integer)
    m_K = new_K
    PropertyChanged "K"
    Draw
End Property

'ĵ�i! ���Ų����έק�H�U�����Ѧ�!
'MemberInfo=7,0,0,0
Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    I.Left = (m_Value / 100) * P.ScaleWidth
    PropertyChanged "Value"
End Property

'��l�ƨϥΪ̱�����ݩ�
Private Sub UserControl_InitProperties()
    m_C = m_def_C
    m_M = m_def_M
    m_Y = m_def_Y
    m_K = m_def_K
    m_Value = m_def_Value
End Sub

'���x�s�ϸ��J�ݩʭ�
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_C = PropBag.ReadProperty("C", m_def_C)
    m_M = PropBag.ReadProperty("M", m_def_M)
    m_Y = PropBag.ReadProperty("Y", m_def_Y)
    m_K = PropBag.ReadProperty("K", m_def_K)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
End Sub

Private Sub UserControl_Resize()
    P.Width = UserControl.ScaleWidth - 6
    P.Height = UserControl.ScaleHeight - 6
    I.Left = 0
    I.Top = P.Height
End Sub

Private Sub UserControl_Show()
    P.ToolTipText = Extender.ToolTipText
    I.ToolTipText = Extender.ToolTipText
End Sub

'�N�ݩʭȼg�^�x�s��
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

'ĵ�i! ���Ų����έק�H�U�����Ѧ�!
'MappingInfo=P,P,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "�Ǧ^�������������N�X (�� Microsoft Windows �Ҵ���)�C"
    hWnd = P.hWnd
End Property

