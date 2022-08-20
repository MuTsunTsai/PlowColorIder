VERSION 5.00
Begin VB.UserControl PlowHand 
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   ScaleHeight     =   145
   ScaleMode       =   3  '像素
   ScaleWidth      =   153
   Begin 北斗色彩識別器.PlowColorB HIColor 
      Height          =   615
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "快速顏色選擇盤"
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
   End
   Begin VB.OptionButton goDec 
      Caption         =   "10進位"
      Height          =   255
      Left            =   570
      TabIndex        =   9
      ToolTipText     =   "選擇採用 10 進位表示"
      Top             =   1185
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton goHex 
      Caption         =   "16進位"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      ToolTipText     =   "選擇採用 16 進位表示"
      Top             =   1185
      Width           =   855
   End
   Begin 北斗色彩識別器.PlowRGB R 
      Height          =   305
      Left            =   360
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "以拖曳方式改變紅色分量"
      Top             =   0
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   529
   End
   Begin 北斗色彩識別器.PlowRGB G 
      Height          =   305
      Left            =   360
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "以拖曳方式改變綠色分量"
      Top             =   390
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   529
   End
   Begin 北斗色彩識別器.PlowRGB B 
      Height          =   300
      Left            =   360
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "以拖曳方式改變藍色分量"
      Top             =   780
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   529
   End
   Begin VB.TextBox SR 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "手動輸入紅色分量"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox SG 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "手動輸入綠色分量"
      Top             =   390
      Width           =   375
   End
   Begin VB.TextBox SB 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "手動輸入藍色分量"
      Top             =   780
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "進制："
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1230
      Width           =   615
   End
   Begin VB.Label SL 
      Caption         =   "紅："
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.Label SL 
      Caption         =   "綠："
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   390
      Width           =   375
   End
   Begin VB.Label SL 
      Caption         =   "藍："
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   780
      Width           =   375
   End
End
Attribute VB_Name = "PlowHand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Const m_def_Color = 0

Private m_Color As Long

Private RO As String, BO As String, GO As String, C As Boolean
Private EventRaised As Boolean

Event Change()

Public Sub SetTextFocus(ByVal Index As Integer)
    If Index = 1 Then
        SR.SetFocus
    Else
        If DecMode Then goDec.SetFocus Else goHex.SetFocus
    End If
End Sub

Private Sub SetColor(ByVal New_Color As Long)
    Static new_R As Integer, new_G As Integer, new_B As Integer
    If New_Color > -1 Then
        m_Color = New_Color
        new_R = m_Color And &HFF
        new_G = (m_Color \ &H100) And &HFF
        new_B = (m_Color \ &H10000) And &HFF
        R.Setvalue 256, new_G, new_B: R.Value = new_R
        G.Setvalue new_R, 256, new_B: G.Value = new_G
        B.Setvalue new_R, new_G, 256: B.Value = new_B
        C = False
        If DecMode Then SR.Text = R.Value Else SR.Text = Hex(R.Value)
        If DecMode Then SG.Text = G.Value Else SG.Text = Hex(G.Value)
        If DecMode Then SB.Text = B.Value Else SB.Text = Hex(B.Value)
        C = True
        RaiseEvent Change
        PropertyChanged "Color"
    End If
End Sub

Private Sub Setvalue()
    If DecMode Then
        SR.Text = R.Value: SG.Text = G.Value: SB.Text = B.Value
    Else
        SR.Text = Hex(R.Value): SG.Text = Hex(G.Value): SB.Text = Hex(B.Value)
    End If
End Sub

Private Sub ChangeColor()
    m_Color = RGB(R.Value, G.Value, B.Value)
    RaiseEvent Change
End Sub

Private Function CheckNumber(ByVal TT As String) As String
    If DecMode Then
        CheckNumber = IIf(IsNumeric(TT), ValueInRange(Val(TT), 0, 255), "-1")
    Else
        CheckNumber = IIf(IsNumeric("&H" & TT), Hex(ValueInRange(Val("&H" & TT), 0, 255)), "-1")
    End If
End Function

Private Sub B_Change(Button As Integer, Shift As Integer, X As Single, Y As Single)
    C = False: R.B = B.Value: G.B = B.Value
    If DecMode Then SB.Text = B.Value Else SB.Text = Hex(B.Value)
    ChangeColor
    C = True
End Sub

Private Sub G_Change(Button As Integer, Shift As Integer, X As Single, Y As Single)
    C = False: R.G = G.Value: B.G = G.Value
    If DecMode Then SG.Text = G.Value Else SG.Text = Hex(G.Value)
    ChangeColor
    C = True
End Sub

Private Sub goDec_GotFocus()
    KBDOptionFocused = 1
End Sub

Private Sub goHex_GotFocus()
    KBDOptionFocused = 2
End Sub

Private Sub goHex_LostFocus()
    KBDOptionFocused = 0
End Sub

Private Sub goDec_LostFocus()
    KBDOptionFocused = 0
End Sub

Private Sub HIColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then SetColor HIColor.Color
End Sub

Private Sub HIColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then SetColor HIColor.Color
End Sub

Private Sub R_Change(Button As Integer, Shift As Integer, X As Single, Y As Single)
    C = False: G.R = R.Value: B.R = R.Value
    If DecMode Then SR.Text = R.Value Else SR.Text = Hex(R.Value)
    ChangeColor
    C = True
End Sub

Private Sub SB_Change()
    Static Old As Integer
    If EventRaised Then Exit Sub
    Old = SB.SelStart: EventRaised = True
    SB.Text = UCase(SB.Text)
    Do While Left(SB.Text, 1) = "0"
        SB.Text = Right(SB.Text, Len(SB.Text) - 1)
        If Old = 1 Then Old = 0
    Loop
    If SB.Text = "" Then SB.Text = 0
    If CheckNumber(SB.Text) = "-1" Then
        SB.Text = BO
    Else
        SB.Text = CheckNumber(SB.Text)
        BO = SB.Text
    End If
    If C Then
        If DecMode Then B.Value = Val(SB.Text) Else B.Value = Val("&H" & SB.Text)
        R.B = B.Value: G.B = B.Value
    End If
    ChangeColor
    If SB.SelStart <> Old Then SB.SelStart = Old
    If SB.Text = "0" Then SB.SelLength = 1
    EventRaised = False
End Sub

Private Sub SB_GotFocus()
    KBDHooked = True
    SB.SelStart = 0
    SB.SelLength = Len(SB.Text)
End Sub

Private Sub SB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If DecMode Then goDec.SetFocus Else goHex.SetFocus
    If KeyCode = vbKeyUp Then SG.SetFocus
    If (Shift And vbCtrlMask) And KeyCode = vbKeyA Then
        SB.SelStart = 0
        SB.SelLength = Len(SB.Text)
    End If
End Sub

Private Sub SB_KeyPress(KeyAscii As Integer)
    If (GetKeyState(vbKeyA) And &H1000) And (GetKeyState(vbKeyControl) And &H1000) Then KeyAscii = 0
End Sub

Private Sub SB_LostFocus()
    KBDHooked = False
End Sub

Private Sub SG_Change()
    Static Old As Integer
    If EventRaised Then Exit Sub
    Old = SG.SelStart: EventRaised = True
    SG.Text = UCase(SG.Text)
    Do While Left(SG.Text, 1) = "0"
        SG.Text = Right(SG.Text, Len(SG.Text) - 1)
        If Old = 1 Then Old = 0
    Loop
    If SG.Text = "" Then SG.Text = 0
    If CheckNumber(SG.Text) = "-1" Then
        SG.Text = GO
    Else
        SG.Text = CheckNumber(SG.Text)
        GO = SG.Text
    End If
    If C Then
        If DecMode Then G.Value = Val(SG.Text) Else G.Value = Val("&H" & SG.Text)
        R.G = G.Value: B.G = G.Value
    End If
    ChangeColor
    If SG.SelStart <> Old Then SG.SelStart = Old
    If SG.Text = "0" Then SG.SelLength = 1
    EventRaised = False
End Sub

Private Sub SG_GotFocus()
    KBDHooked = True
    SG.SelStart = 0
    SG.SelLength = Len(SG.Text)
End Sub

Private Sub SG_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SB.SetFocus
    If KeyCode = vbKeyUp Then SR.SetFocus
    If (Shift And vbCtrlMask) And KeyCode = vbKeyA Then SG.SelStart = 0: SG.SelLength = Len(SG.Text)
End Sub

Private Sub SG_KeyPress(KeyAscii As Integer)
    If (GetKeyState(vbKeyA) And &H1000) And (GetKeyState(vbKeyControl) And &H1000) Then KeyAscii = 0
End Sub

Private Sub SG_LostFocus()
    KBDHooked = False
End Sub

Private Sub SR_Change()
    Static Old As Integer
    If EventRaised Then Exit Sub
    Old = SR.SelStart: EventRaised = True
    SR.Text = UCase(SR.Text)
    Do While Left(SR.Text, 1) = "0"
        SR.Text = Right(SR.Text, Len(SR.Text) - 1)
        If Old = 1 Then Old = 0
    Loop
    If SR.Text = "" Then SR.Text = 0
    If CheckNumber(SR.Text) = "-1" Then
        SR.Text = RO
    Else
        SR.Text = CheckNumber(SR.Text)
        RO = SR.Text
    End If
    If C Then
        If DecMode Then R.Value = Val(SR.Text) Else R.Value = Val("&H" & SR.Text)
        G.R = R.Value: B.R = R.Value
    End If
    ChangeColor
    If SR.SelStart <> Old Then SR.SelStart = Old
    If SR.Text = "0" Then SR.SelLength = 1
    EventRaised = False
End Sub

Private Sub SR_GotFocus()
    KBDHooked = True
    SR.SelStart = 0
    SR.SelLength = Len(SR.Text)
End Sub

Private Sub SR_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SG.SetFocus
    If KeyCode = vbKeyUp Then If DecMode Then goDec.SetFocus Else goHex.SetFocus
    If (Shift And vbCtrlMask) And KeyCode = vbKeyA Then SR.SelStart = 0: SR.SelLength = Len(SR.Text)
End Sub

Private Sub SR_KeyPress(KeyAscii As Integer)
    If (GetKeyState(vbKeyA) And &H1000) And (GetKeyState(vbKeyControl) And &H1000) Then KeyAscii = 0
End Sub

Private Sub SR_LostFocus()
    KBDHooked = False
End Sub

Private Sub UserControl_Initialize()
    R.Setvalue 256, 0, 0
    G.Setvalue 0, 256, 0
    B.Setvalue 0, 0, 256
    RO = 0: GO = 0: BO = 0
    DecMode = True: C = True: EventRaised = False
    HIColor.Draw
End Sub

Private Sub goHex_Click()
    If DecMode Then DecMode = False
    SR.MaxLength = 2: SG.MaxLength = 2: SB.MaxLength = 2
    Setvalue
End Sub

Private Sub goDec_Click()
    If Not DecMode Then DecMode = True
    SR.MaxLength = 3: SG.MaxLength = 3: SB.MaxLength = 3
    Setvalue
End Sub

Public Property Get Color() As Long
    Color = m_Color
End Property

Public Property Let Color(ByVal New_Color As Long)
    SetColor New_Color
End Property

Private Sub UserControl_InitProperties()
    m_Color = m_def_Color
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Color = PropBag.ReadProperty("Color", m_def_Color)
End Sub

Private Sub UserControl_Resize()
    UserControl.Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color", m_Color, m_def_Color)
End Sub

Public Sub Scroll(ByVal Value As Integer)
    Static tmpCURPOS As PointAPI, curHWND
    GetCursorPos tmpCURPOS
    curHWND = WindowFromPoint(tmpCURPOS.X, tmpCURPOS.Y)
    
    If curHWND = SR.hWnd Or curHWND = R.hWnd Then ScrollTxt SR, Value
    If curHWND = SG.hWnd Or curHWND = G.hWnd Then ScrollTxt SG, Value
    If curHWND = SB.hWnd Or curHWND = B.hWnd Then ScrollTxt SB, Value
End Sub

Private Sub ScrollTxt(ByRef txt As Object, ByVal Value As Integer)
    If DecMode Then txt.Text = ValueInRange(Val(txt.Text) + Value, 0, 255) _
        Else txt.Text = Hex(ValueInRange(Val("&H" & txt.Text) + Value, 0, 255))
    txt.SetFocus: txt.SelStart = 0: txt.SelLength = Len(txt.Text)
End Sub




Public Sub OptionKeyDown(ByVal KeyCode As Integer)
    If KeyCode = vbKeyUp Then SB.SetFocus
    If KeyCode = vbKeyDown Then SR.SetFocus
    If KBDOptionFocused = 1 Then
        If KeyCode = vbKeyRight Then goHex.SetFocus
    Else
        If KeyCode = vbKeyLeft Then goDec.SetFocus
    End If
End Sub

