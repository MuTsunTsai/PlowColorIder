VERSION 5.00
Begin VB.UserControl PlowFlatButton 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   480
   End
   Begin VB.Label Cap 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Caption"
      Height          =   180
      Left            =   30
      TabIndex        =   0
      Top             =   30
      UseMnemonic     =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "PlowFlatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private mDown As Boolean, mMove As Boolean

Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOut()

Private Sub Cap_Click()
    UserControl_Click
End Sub

Private Sub Cap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub Cap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Cap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub Timer_Timer()
    Static tmpCURPOS As PointAPI, curHWND
    GetCursorPos tmpCURPOS
    curHWND = WindowFromPoint(tmpCURPOS.X, tmpCURPOS.Y)
    If curHWND <> UserControl.hWnd Then
        mMove = False: mDown = False
        UserControl.Refresh
        RaiseEvent MouseOut
        Timer.Enabled = False
    End If
End Sub

Private Sub UserControl_Paint()
    Static W, H
    W = UserControl.ScaleWidth - 15: H = UserControl.ScaleHeight - 15
    If mDown Then
        Line (0, 0)-(W, 0), vbButtonShadow
        Line (0, 0)-(0, H), vbButtonShadow
        Line (W, 0)-(W, H), vb3DHighlight
        Line (0, H)-(W + 15, H), vb3DHighlight
    ElseIf Not Ambient.UserMode Or mMove Then
        Line (0, 0)-(W, 0), vb3DHighlight
        Line (0, 0)-(0, H), vb3DHighlight
        Line (W, 0)-(W, H), vbButtonShadow
        Line (0, H)-(W + 15, H), vbButtonShadow
    End If
End Sub

Private Sub UserControl_Resize()
    ' 下一行解決在舊版的 VB6 Runtime 環境中的字型問題
    Cap.Font.Size = UserControl.Font.Size
    UserControl.Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mDown = True: UserControl.Refresh
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static T As PointAPI, curHWND
    If Not mMove Then
        If X >= 0 And Y >= 0 And X < UserControl.ScaleWidth And Y < UserControl.ScaleHeight Then
            mMove = True: Timer.Enabled = True
            UserControl.Refresh
            If Button = 1 Then UserControl_MouseDown Button, Shift, X, Y
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 0 And Y >= 0 And X < UserControl.ScaleWidth And Y < UserControl.ScaleHeight Then
        If Button = 1 Then mDown = False: UserControl.Refresh
        RaiseEvent MouseUp(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Cap.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Cap.Caption = PropBag.ReadProperty("Caption", "Caption")
    Cap.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_Show()
    Cap.ToolTipText = Extender.ToolTipText
End Sub

Public Sub SetToolTipText(ByVal New_ToolTipText As String)
    Cap.ToolTipText = New_ToolTipText
    Extender.ToolTipText = New_ToolTipText
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ForeColor", Cap.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Caption", Cap.Caption, "Caption")
    Call PropBag.WriteProperty("Font", Cap.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "傳回或設定在物件中，顯示文字與圖形的前景色彩。"
    ForeColor = Cap.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Cap.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "傳回或設定顯示於物件標題列或圖示下方的文字。"
    Caption = Cap.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Cap.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'警告! 切勿移除或修改以下的註解行!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "傳回或設定在物件中，顯示文字與圖形的背景色彩。"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'警告! 切勿移除或修改以下的註解行!
'MappingInfo=Cap,Cap,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "傳回一個 Font 物件"
Attribute Font.VB_UserMemId = -512
    Set Font = Cap.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Cap.Font = New_Font
    PropertyChanged "Font"
End Property

