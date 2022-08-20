VERSION 5.00
Begin VB.UserControl PlowNumber 
   BackColor       =   &H80000005&
   BorderStyle     =   1  '單線固定
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  '像素
   ScaleWidth      =   320
   Begin VB.TextBox TB 
      BorderStyle     =   0  '沒有框線
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
      Left            =   15
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "手動輸入藍色分量"
      Top             =   15
      Width           =   735
   End
   Begin VB.VScrollBar VS 
      Height          =   255
      Left            =   720
      Max             =   255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   180
   End
End
Attribute VB_Name = "PlowNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Event Declarations:
Event Change()
Attribute Change.VB_Description = "發生於控制項的內容改變時。"

'警告! 切勿移除或修改以下的註解行!
'MappingInfo=VS,VS,-1,Max
Public Property Get Max() As Integer
Attribute Max.VB_Description = "傳回或設定代表捲軸位置的 Value 屬性值上限。"
    Max = Vs.Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    Vs.Max() = New_Max
    PropertyChanged "Max"
End Property

'警告! 切勿移除或修改以下的註解行!
'MappingInfo=VS,VS,-1,Min
Public Property Get Min() As Integer
Attribute Min.VB_Description = "傳回或設定代表捲軸位置的 Value 屬性值上限。"
    Min = Vs.Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    Vs.Min() = New_Min
    PropertyChanged "Min"
End Property

'警告! 切勿移除或修改以下的註解行!
'MappingInfo=VS,VS,-1,Value
Public Property Get Value() As Integer
Attribute Value.VB_Description = "傳回或設定一個物件的值。"
    Value = Vs.Max - Vs.Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    Vs.Value() = Vs.Max - New_Value
    PropertyChanged "Value"
End Property

'由儲存區載入屬性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Vs.Max = PropBag.ReadProperty("Max", 255)
    Vs.Min = PropBag.ReadProperty("Min", 0)
    Vs.Value = PropBag.ReadProperty("Value", 255)
End Sub

Private Sub UserControl_Resize()
    TB.Height = UserControl.ScaleHeight - 1
    TB.Width = UserControl.ScaleWidth - 13
    Vs.Left = TB.Width + 1
    Vs.Height = TB.Height + 1
End Sub

'將屬性值寫回儲存區
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Max", Vs.Max, 255)
    Call PropBag.WriteProperty("Min", Vs.Min, 0)
    Call PropBag.WriteProperty("Value", Vs.Value, 0)
End Sub

Private Sub Vs_Change()
    TB.Text = Vs.Max - Vs.Value
    RaiseEvent Change
End Sub
