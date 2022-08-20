VERSION 5.00
Begin VB.UserControl Weblink 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label lblWeblink 
      Caption         =   "Weblink Control"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Weblink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Default Property Values:
Private Const m_def_Location = ""

'Property Variables:
Private m_Location As String

'Event Declarations:
Event Click() 'MappingInfo=lblWeblink,lblWeblink,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=lblWeblink,lblWeblink,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblWeblink,lblWeblink,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblWeblink,lblWeblink,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblWeblink,lblWeblink,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Public Function HyperJump(ByVal URL As String) As Long
    On Error Resume Next
    HyperJump = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblWeblink,lblWeblink,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = lblWeblink.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lblWeblink.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblWeblink,lblWeblink,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblWeblink.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblWeblink.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblWeblink,lblWeblink,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = lblWeblink.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    lblWeblink.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblWeblink,lblWeblink,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblWeblink.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblWeblink.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblWeblink,lblWeblink,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = lblWeblink.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    lblWeblink.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblWeblink,lblWeblink,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = lblWeblink.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    lblWeblink.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblWeblink,lblWeblink,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    lblWeblink.Refresh
End Sub

Private Sub lblWeblink_Click()
HyperJump (m_Location)
End Sub

Private Sub lblWeblink_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub lblWeblink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor 65581
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblWeblink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor 65581
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblWeblink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor 65581
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get Location() As String
Attribute Location.VB_Description = "Weblink Location can be Internet or Local Resource depending on protocol"
    Location = m_Location
End Property

Public Property Let Location(ByVal New_Location As String)
    m_Location = New_Location
    PropertyChanged "Location"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Location = m_def_Location
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblWeblink.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblWeblink.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    lblWeblink.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblWeblink.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    lblWeblink.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_Location = PropBag.ReadProperty("Location", m_def_Location)
    lblWeblink.Caption = PropBag.ReadProperty("Caption", "Weblink Control")
    lblWeblink.Alignment = PropBag.ReadProperty("Alignment", 0)
End Sub

Private Sub UserControl_Resize()
' Move to 0,0 and resize lblWeblink to control size
lblWeblink.Top = 0
lblWeblink.Left = 0
lblWeblink.Width = UserControl.Width
lblWeblink.Height = UserControl.Height
lblWeblink.FontUnderline = True
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", lblWeblink.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", lblWeblink.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", lblWeblink.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", lblWeblink.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", lblWeblink.BorderStyle, 0)
    Call PropBag.WriteProperty("Location", m_Location, m_def_Location)
    Call PropBag.WriteProperty("Caption", lblWeblink.Caption, "Weblink Control")
    Call PropBag.WriteProperty("Alignment", lblWeblink.Alignment, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblWeblink,lblWeblink,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblWeblink.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblWeblink.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblWeblink,lblWeblink,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = lblWeblink.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    lblWeblink.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

