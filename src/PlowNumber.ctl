VERSION 5.00
Begin VB.UserControl PlowNumber 
   BackColor       =   &H80000005&
   BorderStyle     =   1  '��u�T�w
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  '����
   ScaleWidth      =   320
   Begin VB.TextBox TB 
      BorderStyle     =   0  '�S���ؽu
      BeginProperty Font 
         Name            =   "�ө���"
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
      ToolTipText     =   "��ʿ�J�Ŧ���q"
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
Attribute Change.VB_Description = "�o�ͩ󱱨�����e���ܮɡC"

'ĵ�i! ���Ų����έק�H�U�����Ѧ�!
'MappingInfo=VS,VS,-1,Max
Public Property Get Max() As Integer
Attribute Max.VB_Description = "�Ǧ^�γ]�w�N���b��m�� Value �ݩʭȤW���C"
    Max = Vs.Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    Vs.Max() = New_Max
    PropertyChanged "Max"
End Property

'ĵ�i! ���Ų����έק�H�U�����Ѧ�!
'MappingInfo=VS,VS,-1,Min
Public Property Get Min() As Integer
Attribute Min.VB_Description = "�Ǧ^�γ]�w�N���b��m�� Value �ݩʭȤW���C"
    Min = Vs.Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    Vs.Min() = New_Min
    PropertyChanged "Min"
End Property

'ĵ�i! ���Ų����έק�H�U�����Ѧ�!
'MappingInfo=VS,VS,-1,Value
Public Property Get Value() As Integer
Attribute Value.VB_Description = "�Ǧ^�γ]�w�@�Ӫ��󪺭ȡC"
    Value = Vs.Max - Vs.Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    Vs.Value() = Vs.Max - New_Value
    PropertyChanged "Value"
End Property

'���x�s�ϸ��J�ݩʭ�
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

'�N�ݩʭȼg�^�x�s��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Max", Vs.Max, 255)
    Call PropBag.WriteProperty("Min", Vs.Min, 0)
    Call PropBag.WriteProperty("Value", Vs.Value, 0)
End Sub

Private Sub Vs_Change()
    TB.Text = Vs.Max - Vs.Value
    RaiseEvent Change
End Sub
