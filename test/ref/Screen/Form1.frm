VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   2655
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   840
      Top             =   120
   End
   Begin VB.PictureBox picScr 
      AutoRedraw      =   -1  'True
      Height          =   1815
      Left            =   120
      ScaleHeight     =   117
      ScaleMode       =   3  '像素
      ScaleWidth      =   157
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開始"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private mouseX As Integer, mouseY As Integer

Private Sub Command1_Click()
    EnableMSHook
    picScr.BackColor = vbBlack
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2: Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Public Sub Scr(ByVal X As Integer, ByVal Y As Integer)
    'If X < 9 Then X = 9
    'If Y < 7 Then Y = 7
    'If X > Me.ScaleWidth - 10 Then X = Me.ScaleWidth - 10
    'If Y > Me.ScaleHeight - 8 Then Y = Me.ScaleHeight - 8
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    If X >= Screen.Width / Screen.TwipsPerPixelX Then X = Screen.Width / Screen.TwipsPerPixelX - 1
    If Y >= Screen.Height / Screen.TwipsPerPixelY Then Y = Screen.Height / Screen.TwipsPerPixelY - 1
    mouseX = X: mouseY = Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHookMS
End Sub

Private Sub Timer1_Timer()
    Dim hWndSrc As Long
    Dim hDCSrc As Long
    
    hWndSrc = GetDesktopWindow()
    hDCSrc = GetDC(hWndSrc)

    picScr.Cls
    StretchBlt picScr.hDC, 0, 0, picScr.ScaleWidth, picScr.ScaleHeight, _
        hDCSrc, mouseX - 9, mouseY - 7, 19, 15, vbSrcCopy
    picScr.Refresh
End Sub
