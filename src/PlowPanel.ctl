VERSION 5.00
Begin VB.UserControl PlowPanel 
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1260
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   81
   ScaleMode       =   3  '像素
   ScaleWidth      =   84
End
Attribute VB_Name = "PlowPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "發生於使用者移動滑鼠時。"

Private Sub UserControl_Paint()
    Static W, H
    W = UserControl.ScaleWidth - 1: H = UserControl.ScaleHeight - 1
    Line (0, 0)-(W, 0), vb3DHighlight
    Line (0, 0)-(0, H), vb3DHighlight
    Line (W, 0)-(W, H), vbButtonShadow
    Line (0, H)-(W + 1, H), vbButtonShadow
End Sub

Private Sub UserControl_Resize()
    UserControl.Refresh
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
