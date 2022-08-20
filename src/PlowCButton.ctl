VERSION 5.00
Begin VB.UserControl PlowCButton 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Pic 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   1455
      Left            =   15
      ScaleHeight     =   1455
      ScaleWidth      =   1695
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   1695
   End
End
Attribute VB_Name = "PlowCButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private mDown As Boolean

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub Pic_Click()
    UserControl_Click
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_Paint()
    Static W, H
    W = UserControl.ScaleWidth - 15: H = UserControl.ScaleHeight - 15
    Pic.Move 15, 15, W - 15, H - 15
    If mDown Then
        Line (0, 0)-(W, 0), vbButtonShadow
        Line (0, 0)-(0, H), vbButtonShadow
        Line (W, 0)-(W, H), vb3DHighlight
        Line (0, H)-(W + 15, H), vb3DHighlight
    Else
        Line (0, 0)-(W, 0), vb3DHighlight
        Line (0, 0)-(0, H), vb3DHighlight
        Line (W, 0)-(W, H), vbButtonShadow
        Line (0, H)-(W + 15, H), vbButtonShadow
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Refresh
End Sub

Public Property Get BackColor() As Long
    BackColor = Pic.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    Pic.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mDown = True: UserControl.Refresh
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mDown = False: UserControl.Refresh
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Pic.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Pic.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_Show()
    Pic.ToolTipText = Extender.ToolTipText
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", Pic.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackColor", Pic.BackColor, &H8000000F)
End Sub
