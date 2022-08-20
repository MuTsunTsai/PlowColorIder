VERSION 5.00
Begin VB.UserControl PlowCHand 
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   ScaleHeight     =   145
   ScaleMode       =   3  '像素
   ScaleWidth      =   153
   Begin VB.TextBox TK 
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
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "手動輸入黑色分量"
      Top             =   1170
      Width           =   375
   End
   Begin VB.TextBox TY 
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
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "手動輸入黃色分量"
      Top             =   780
      Width           =   375
   End
   Begin VB.TextBox TM 
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
      TabIndex        =   5
      Text            =   "0"
      ToolTipText     =   "手動輸入洋紅分量"
      Top             =   390
      Width           =   375
   End
   Begin VB.TextBox TC 
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
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "手動輸入青色分量"
      Top             =   0
      Width           =   375
   End
   Begin 北斗色彩識別器.PlowCMYK PC 
      Height          =   300
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "以拖曳方式改變青色分量"
      Top             =   0
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   529
   End
   Begin 北斗色彩識別器.PlowCMYK PM 
      Height          =   300
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "以拖曳方式改變洋紅分量"
      Top             =   390
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   529
   End
   Begin 北斗色彩識別器.PlowCMYK PY 
      Height          =   300
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "以拖曳方式改變黃色分量"
      Top             =   780
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   529
   End
   Begin 北斗色彩識別器.PlowCMYK PK 
      Height          =   300
      Left            =   360
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "以拖曳方式改變黑色分量"
      Top             =   1170
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   529
   End
   Begin 北斗色彩識別器.PlowColorB HIColor 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "快速顏色選擇盤"
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
   End
   Begin VB.Label LK 
      Caption         =   "黑："
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1170
      Width           =   375
   End
   Begin VB.Label LY 
      Caption         =   "黃："
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   780
      Width           =   375
   End
   Begin VB.Label LM 
      Caption         =   "紅："
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   390
      Width           =   375
   End
   Begin VB.Label LC 
      Caption         =   "青："
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "PlowCHand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Const m_def_Color = 0

Private Changing As Boolean, EventRaised As Boolean
Private CO As String, MO As String, YO As String, KO As String
Private CC As New ColorCMYK

Event Change()

Public Sub SetTextFocus(ByVal Index As Integer)
    If Index = 1 Then TC.SetFocus Else TK.SetFocus
End Sub

Private Sub SetColor(ByVal New_Color As Long)
    Static new_R As Integer, new_G As Integer, new_B As Integer
    If New_Color > -1 Then
        CC.Color = New_Color
        PC.Setvalue 101, CC.m, CC.Y, CC.K: PC.Value = CC.C
        PM.Setvalue CC.C, 101, CC.Y, CC.K: PM.Value = CC.m
        PY.Setvalue CC.C, CC.m, 101, CC.K: PY.Value = CC.Y
        PK.Setvalue CC.C, CC.m, CC.Y, 101: PK.Value = CC.K
        Changing = True
        TC.Text = PC.Value
        TM.Text = PM.Value
        TY.Text = PY.Value
        TK.Text = PK.Value
        Changing = False
        RaiseEvent Change
        PropertyChanged "Color"
    End If
End Sub

Private Function CheckNnumber(ByVal TT As String) As Integer
    CheckNnumber = IIf(IsNumeric(TT), ValueInRange(Val(TT), 0, 100), -1)
End Function

Private Sub HIColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then SetColor HIColor.Color
End Sub

Private Sub HIColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then SetColor HIColor.Color
End Sub

Private Sub PC_Change(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Changing = True
    PM.C = PC.Value: PY.C = PC.Value: PK.C = PC.Value
    TC.Text = PC.Value
    CC.C = PC.Value
    RaiseEvent Change
    Changing = False
End Sub

Private Sub PK_Change(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Changing = True
    PC.K = PK.Value: PM.K = PK.Value: PY.K = PK.Value
    TK.Text = PK.Value
    CC.K = PK.Value
    RaiseEvent Change
    Changing = False
End Sub

Private Sub PM_Change(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Changing = True
    PC.m = PM.Value: PY.m = PM.Value: PK.m = PM.Value
    TM.Text = PM.Value
    CC.m = PM.Value
    RaiseEvent Change
    Changing = False
End Sub

Private Sub PY_Change(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Changing = True
    PC.Y = PY.Value: PM.Y = PY.Value: PK.Y = PY.Value
    TY.Text = PY.Value
    CC.Y = PY.Value
    RaiseEvent Change
    Changing = False
End Sub

Private Sub TC_Change()
    Static Old As Integer
    If EventRaised Then Exit Sub
    Old = TC.SelStart: EventRaised = True
    Do While Left(TC.Text, 1) = "0"
        TC.Text = Right(TC.Text, Len(TC.Text) - 1)
        If Old = 1 Then Old = 0
    Loop
    If TC.Text = "" Then TC.Text = 0
    If CheckNnumber(TC.Text) = -1 Then
        TC.Text = CO
    Else
        TC.Text = CheckNnumber(TC.Text)
        CO = TC.Text
    End If
    If Not Changing Then
        PC.Value = Val(TC.Text)
        PM.C = PC.Value: PY.C = PC.Value: PK.C = PC.Value
        CC.C = PC.Value
        RaiseEvent Change
    End If
    If TC.SelStart <> Old Then TC.SelStart = Old
    If TC.Text = "0" Then TC.SelLength = 1
    EventRaised = False
End Sub

Private Sub TC_GotFocus()
    KBDHooked = True
    TC.SelStart = 0
    TC.SelLength = Len(TC.Text)
End Sub

Private Sub TC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then TM.SetFocus
    If KeyCode = vbKeyUp Then TK.SetFocus
    If (Shift And vbCtrlMask) And KeyCode = vbKeyA Then TC.SelStart = 0: TC.SelLength = Len(TC.Text)
End Sub

Private Sub TC_KeyPress(KeyAscii As Integer)
    If (GetKeyState(vbKeyA) And &H1000) And (GetKeyState(vbKeyControl) And &H1000) Then KeyAscii = 0
End Sub

Private Sub TC_LostFocus()
    KBDHooked = False
End Sub

Private Sub TK_Change()
    Static Old As Integer
    If EventRaised Then Exit Sub
    Old = TK.SelStart: EventRaised = True
    Do While Left(TK.Text, 1) = "0"
        TK.Text = Right(TK.Text, Len(TK.Text) - 1)
        If Old = 1 Then Old = 0
    Loop
    If TK.Text = "" Then TK.Text = 0
    If CheckNnumber(TK.Text) = -1 Then
        TK.Text = KO
    Else
        TK.Text = CheckNnumber(TK.Text)
        KO = TK.Text
    End If
    If Not Changing Then
        PK.Value = Val(TK.Text)
        PC.K = PK.Value: PM.K = PK.Value: PY.K = PK.Value
        CC.K = PK.Value
        RaiseEvent Change
    End If
    If TK.SelStart <> Old Then TK.SelStart = Old
    If TK.Text = "0" Then TK.SelLength = 1
    EventRaised = False
End Sub

Private Sub TK_GotFocus()
    KBDHooked = True
    TK.SelStart = 0
    TK.SelLength = Len(TK.Text)
End Sub

Private Sub TK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then TC.SetFocus
    If KeyCode = vbKeyUp Then TY.SetFocus
    If (Shift And vbCtrlMask) And KeyCode = vbKeyA Then TK.SelStart = 0: TK.SelLength = Len(TK.Text)
End Sub

Private Sub TK_KeyPress(KeyAscii As Integer)
    If (GetKeyState(vbKeyA) And &H1000) And (GetKeyState(vbKeyControl) And &H1000) Then KeyAscii = 0
End Sub

Private Sub TK_LostFocus()
    KBDHooked = False
End Sub

Private Sub TM_Change()
    Static Old As Integer
    If EventRaised Then Exit Sub
    Old = TM.SelStart: EventRaised = True
    Do While Left(TM.Text, 1) = "0"
        TM.Text = Right(TM.Text, Len(TM.Text) - 1)
        If Old = 1 Then Old = 0
    Loop
    If TM.Text = "" Then TM.Text = 0
    If CheckNnumber(TM.Text) = -1 Then
        TM.Text = MO
    Else
        TM.Text = CheckNnumber(TM.Text)
        MO = TM.Text
    End If
    If Not Changing Then
        PM.Value = Val(TM.Text)
        PC.m = PM.Value: PY.m = PM.Value: PK.m = PM.Value
        CC.m = PM.Value
        RaiseEvent Change
    End If
    If TM.SelStart <> Old Then TM.SelStart = Old
    If TM.Text = "0" Then TM.SelLength = 1
    EventRaised = False
End Sub

Private Sub TM_GotFocus()
    KBDHooked = True
    TM.SelStart = 0
    TM.SelLength = Len(TM.Text)
End Sub

Private Sub TM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then TY.SetFocus
    If KeyCode = vbKeyUp Then TC.SetFocus
    If (Shift And vbCtrlMask) And KeyCode = vbKeyA Then TM.SelStart = 0: TM.SelLength = Len(TM.Text)
End Sub

Private Sub TM_KeyPress(KeyAscii As Integer)
    If (GetKeyState(vbKeyA) And &H1000) And (GetKeyState(vbKeyControl) And &H1000) Then KeyAscii = 0
End Sub

Private Sub Tm_LostFocus()
    KBDHooked = False
End Sub

Private Sub TY_Change()
    Static Old As Integer
    If EventRaised Then Exit Sub
    Old = TY.SelStart: EventRaised = True
    Do While Left(TY.Text, 1) = "0"
        TY.Text = Right(TY.Text, Len(TY.Text) - 1)
        If Old = 1 Then Old = 0
    Loop
    If TY.Text = "" Then TY.Text = 0
    If CheckNnumber(TY.Text) = -1 Then
        TY.Text = YO
    Else
        TY.Text = CheckNnumber(TY.Text)
        YO = TY.Text
    End If
    If Not Changing Then
        PY.Value = Val(TY.Text)
        PC.Y = PY.Value: PM.Y = PY.Value: PK.Y = PY.Value
        CC.Y = PY.Value
        RaiseEvent Change
    End If
    If TY.SelStart <> Old Then TY.SelStart = Old
    If TY.Text = "0" Then TY.SelLength = 1
    EventRaised = False
End Sub

Private Sub TY_GotFocus()
    KBDHooked = True
    TY.SelStart = 0
    TY.SelLength = Len(TY.Text)
End Sub

Private Sub TY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then TK.SetFocus
    If KeyCode = vbKeyUp Then TM.SetFocus
    If (Shift And vbCtrlMask) And KeyCode = vbKeyA Then TY.SelStart = 0: TY.SelLength = Len(TY.Text)
End Sub

Private Sub TY_KeyPress(KeyAscii As Integer)
    If (GetKeyState(vbKeyA) And &H1000) And (GetKeyState(vbKeyControl) And &H1000) Then KeyAscii = 0
End Sub

Private Sub Ty_LostFocus()
    KBDHooked = False
End Sub

Private Sub UserControl_Initialize()
    PC.Setvalue 101, 0, 0, 0
    PM.Setvalue 0, 101, 0, 0
    PY.Setvalue 0, 0, 101, 0
    PK.Setvalue 0, 0, 0, 101
    Changing = False: EventRaised = False
    CO = "0": MO = "0": YO = "0": KO = "0"
    CC.Color = m_def_Color
    HIColor.Draw
End Sub

Public Property Get Color() As Long
    Color = CC.Color
End Property

Public Property Let Color(ByVal New_Color As Long)
    SetColor New_Color
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    CC.Color = PropBag.ReadProperty("Color", m_def_Color)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color", CC.Color, m_def_Color)
End Sub

Public Sub Scroll(ByVal Value As Integer)
    Static tmpCURPOS As PointAPI, curHWND
    GetCursorPos tmpCURPOS
    curHWND = WindowFromPoint(tmpCURPOS.X, tmpCURPOS.Y)
    
    If curHWND = TC.hWnd Or curHWND = PC.hWnd Then ScrollTxt TC, Value
    If curHWND = TM.hWnd Or curHWND = PM.hWnd Then ScrollTxt TM, Value
    If curHWND = TY.hWnd Or curHWND = PY.hWnd Then ScrollTxt TY, Value
    If curHWND = TK.hWnd Or curHWND = PK.hWnd Then ScrollTxt TK, Value
End Sub

Private Sub ScrollTxt(ByRef txt As Object, ByVal Value As Integer)
    txt.Text = ValueInRange(Val(txt.Text) + Value, 0, 100)
    txt.SetFocus: txt.SelStart = 0: txt.SelLength = Len(txt.Text)
End Sub

