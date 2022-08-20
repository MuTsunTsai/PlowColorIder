VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  '單線固定
   Caption         =   "北斗色彩識別器特別版"
   ClientHeight    =   3375
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  '像素
   ScaleWidth      =   377
   StartUpPosition =   2  '螢幕中央
   Begin 北斗色彩識別器SE.PlowPanel PlowPanel3 
      Height          =   1320
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2055
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2328
      Begin VB.PictureBox HI 
         AutoRedraw      =   -1  'True
         Height          =   1095
         Left            =   120
         MousePointer    =   10  '往上指
         ScaleHeight     =   69
         ScaleMode       =   3  '像素
         ScaleWidth      =   357
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         Width           =   5415
      End
   End
   Begin 北斗色彩識別器SE.PlowPanel PlowPanel2 
      Height          =   2040
      Left            =   3735
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   15
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   3598
      Begin VB.PictureBox TColor 
         BackColor       =   &H00000000&
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1755
         ScaleWidth      =   1635
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   1695
      End
   End
   Begin 北斗色彩識別器SE.PlowPanel PlowPanel1 
      Height          =   2040
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3598
      Begin VB.TextBox TK 
         Height          =   270
         Left            =   3000
         TabIndex        =   15
         Text            =   "0"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox TY 
         Height          =   270
         Left            =   3000
         TabIndex        =   14
         Text            =   "0"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox TM 
         Height          =   270
         Left            =   3000
         TabIndex        =   13
         Text            =   "0"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TC 
         Height          =   270
         Left            =   3000
         TabIndex        =   12
         Text            =   "0"
         Top             =   120
         Width           =   495
      End
      Begin 北斗色彩識別器SE.PlowCMYK PM 
         Height          =   360
         Left            =   720
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
      End
      Begin 北斗色彩識別器SE.PlowCMYK PC 
         Height          =   360
         Left            =   720
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
      End
      Begin 北斗色彩識別器SE.PlowCMYK PY 
         Height          =   360
         Left            =   720
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
      End
      Begin 北斗色彩識別器SE.PlowCMYK PK 
         Height          =   360
         Left            =   720
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
      End
      Begin VB.Label LK 
         Caption         =   "黑："
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label LY 
         Caption         =   "黃："
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LM 
         Caption         =   "洋紅："
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label LC 
         Caption         =   "青："
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   376
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu M_HLP 
      Caption         =   "說明(&H)"
      Begin VB.Menu MH_About 
         Caption         =   "關於(&A)"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Changing As Boolean, EventRaised As Boolean
Private CO As String, MO As String, YO As String, KO As String

Private Sub UpdateColor()
    Static CC As New ColorCMYK
    CC.C = PC.Value: CC.M = PM.Value: CC.Y = PY.Value: CC.K = PK.Value
    TColor.BackColor = CC.Color
End Sub

Private Function IsN(ByVal TT As String) As Boolean
    Static I As Integer
    IsN = True
    If Not IsNumeric(TT) Then
        IsN = False
    Else
        TT = Val(TT)
        If TT < 0 Or TT > 100 Then IsN = False
    End If
End Function

Private Sub Form_Load()
    Static C As Object, W As Integer, H As Integer
    Static I As Integer, J As Integer
    If GetDeviceCaps(Me.hdc, 12&) < 24 Then MsgBox "為了正確分析色彩，建議您將螢幕調整至全彩模式。"
    PrevWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WndProc)

    Set C = New ColorHSL: C.S = 100
    W = HI.ScaleWidth - 1: H = HI.ScaleHeight - 1
    For I = 0 To W
        C.H = I / W * 359
        For J = 0 To H
            C.I = (H - J) / H * 100
            SetPixel HI.hdc, I, J, C.Color
        Next J
    Next I
    HI.Refresh

    PC.Setvalue 101, 0, 0, 0
    PM.Setvalue 0, 101, 0, 0
    PY.Setvalue 0, 0, 101, 0
    PK.Setvalue 0, 0, 0, 101
    Changing = False: EventRaised = False
    CO = "0": MO = "0": YO = "0": KO = "0"
End Sub

Private Sub HI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static CC As New ColorCMYK
    If Button = 1 Then
        LockCursor HI.hWnd, 2
        CC.Color = HI.Point(X, Y)
        TC.Text = CInt(CC.C): TM.Text = CInt(CC.M): TY.Text = CInt(CC.Y): TK.Text = CInt(CC.K)
        UpdateColor
    End If
End Sub

Private Sub HI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static CC As New ColorCMYK
    If Button = 1 Then
        LockCursor HI.hWnd, 2
        CC.Color = HI.Point(X, Y)
        TC.Text = CInt(CC.C): TM.Text = CInt(CC.M): TY.Text = CInt(CC.Y): TK.Text = CInt(CC.K)
        UpdateColor
    End If
End Sub

Private Sub HI_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then UnLockCursor
End Sub

Private Sub MH_About_Click()
    About.Show 1
End Sub

Private Sub PC_Change(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Changing = True
    PM.C = PC.Value: PY.C = PC.Value: PK.C = PC.Value
    TC.Text = PC.Value
    UpdateColor
    Changing = False
End Sub

Private Sub PK_Change(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Changing = True
    PC.K = PK.Value: PM.K = PK.Value: PY.K = PK.Value
    TK.Text = PK.Value
    UpdateColor
    Changing = False
End Sub

Private Sub PM_Change(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Changing = True
    PC.M = PM.Value: PY.M = PM.Value: PK.M = PM.Value
    TM.Text = PM.Value
    UpdateColor
    Changing = False
End Sub

Private Sub PY_Change(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Changing = True
    PC.Y = PY.Value: PM.Y = PY.Value: PK.Y = PY.Value
    TY.Text = PY.Value
    UpdateColor
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
    If IsN(TC.Text) Then CO = TC.Text Else TC.Text = CO
    If Not Changing Then
        PC.Value = Val(TC.Text)
        PM.C = PC.Value: PY.C = PC.Value: PK.C = PC.Value
        UpdateColor
    End If
    If TC.SelStart <> Old Then TC.SelStart = Old
    If TC.Text = "0" Then TC.SelLength = 1
    EventRaised = False
End Sub

Private Sub TC_GotFocus()
    EnableKBDHook
    TC.SelStart = 0
    TC.SelLength = Len(TC.Text)
End Sub

Private Sub TC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then TM.SetFocus
    If KeyCode = vbKeyUp Then TK.SetFocus
    If GetAsyncKeyState(vbKeyControl) <> 0 And KeyCode = vbKeyA Then TC.SelStart = 0: TC.SelLength = Len(TC.Text)
End Sub

Private Sub TC_LostFocus()
    UnHookKBD
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
    If IsN(TK.Text) Then KO = TK.Text Else TK.Text = KO
    If Not Changing Then
        PK.Value = Val(TK.Text)
        PC.K = PK.Value: PM.K = PK.Value: PY.K = PK.Value
        UpdateColor
    End If
    If TK.SelStart <> Old Then TK.SelStart = Old
    If TK.Text = "0" Then TK.SelLength = 1
    EventRaised = False
End Sub

Private Sub TK_GotFocus()
    EnableKBDHook
    TK.SelStart = 0
    TK.SelLength = Len(TK.Text)
End Sub

Private Sub TK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then TC.SetFocus
    If KeyCode = vbKeyUp Then TY.SetFocus
    If GetAsyncKeyState(vbKeyControl) <> 0 And KeyCode = vbKeyA Then TK.SelStart = 0: TK.SelLength = Len(TK.Text)
End Sub

Private Sub Tk_LostFocus()
    UnHookKBD
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
    If IsN(TM.Text) Then MO = TM.Text Else TM.Text = MO
    If Not Changing Then
        PM.Value = Val(TM.Text)
        PC.M = PM.Value: PY.M = PM.Value: PK.M = PM.Value
        UpdateColor
    End If
    If TM.SelStart <> Old Then TM.SelStart = Old
    If TM.Text = "0" Then TM.SelLength = 1
    EventRaised = False
End Sub

Private Sub TM_GotFocus()
    EnableKBDHook
    TM.SelStart = 0
    TM.SelLength = Len(TM.Text)
End Sub

Private Sub TM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then TY.SetFocus
    If KeyCode = vbKeyUp Then TC.SetFocus
    If GetAsyncKeyState(vbKeyControl) <> 0 And KeyCode = vbKeyA Then TM.SelStart = 0: TM.SelLength = Len(TM.Text)
End Sub

Private Sub Tm_LostFocus()
    UnHookKBD
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
    If IsN(TY.Text) Then YO = TY.Text Else TY.Text = YO
    If Not Changing Then
        PY.Value = Val(TY.Text)
        PC.Y = PY.Value: PM.Y = PY.Value: PK.Y = PY.Value
        UpdateColor
    End If
    If TY.SelStart <> Old Then TY.SelStart = Old
    If TY.Text = "0" Then TY.SelLength = 1
    EventRaised = False
End Sub

Private Sub TY_GotFocus()
    EnableKBDHook
    TY.SelStart = 0
    TY.SelLength = Len(TY.Text)
End Sub

Private Sub TY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then TK.SetFocus
    If KeyCode = vbKeyUp Then TM.SetFocus
    If GetAsyncKeyState(vbKeyControl) <> 0 And KeyCode = vbKeyA Then TY.SelStart = 0: TY.SelLength = Len(TY.Text)
End Sub

Private Sub Ty_LostFocus()
    UnHookKBD
End Sub
