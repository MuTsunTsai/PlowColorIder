VERSION 5.00
Begin VB.Form FScreen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  '沒有框線
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  '像素
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.PictureBox P 
      Appearance      =   0  '平面
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      MouseIcon       =   "FScreen.frx":0000
      MousePointer    =   99  '自訂
      ScaleHeight     =   89
      ScaleMode       =   3  '像素
      ScaleWidth      =   145
      TabIndex        =   0
      ToolTipText     =   "按一下以選取螢幕上任意點之色彩"
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "FScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const BD = 2

Private sX As Single, Sy As Single
Private CBDC As Long, CBBitmap As Long

Private Sub Form_Load()
    Static I As Integer, j As Integer, K As Integer, R As RECT, B As Long
    frmMain.picScr.MousePointer = 0
    sX = frmMain.picScr.ScaleWidth / 19: Sy = frmMain.picScr.ScaleHeight / 15
    CBDC = CreateCompatibleDC(0)
    CBBitmap = CreateDIBSec(Me.hDC, frmMain.picScr.ScaleWidth, frmMain.picScr.ScaleHeight)
    B = CreateSolidBrush(&HC0C0C0)
    SelectObject CBDC, CBBitmap
    For j = 0 To frmMain.picScr.ScaleHeight Step BD
        If (j / BD) Mod 2 Then K = 0 Else K = BD
        For I = K To frmMain.picScr.ScaleWidth Step BD * 2
            R.Top = j: R.Left = I
            R.Bottom = j + BD: R.Right = I + BD
            FillRect CBDC, R, B
        Next I
    Next j
End Sub

Private Sub P_Click()
    frmMain.GetColorInfo frmMain.picScr.Point(sX * 9, Sy * 7)
    frmMain.ScrCaptureEnd False
End Sub

Private Sub P_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetWindowPos frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
End Sub

Private Sub P_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.lblPos.Caption = "座標：( " & X & " , " & Y & " )"
    BitBlt frmMain.picScr.hDC, 0, 0, frmMain.picScr.ScaleWidth, frmMain.picScr.ScaleHeight, _
        CBDC, 0, 0, vbSrcCopy
    StretchBlt frmMain.picScr.hDC, 0, 0, frmMain.picScr.ScaleWidth, frmMain.picScr.ScaleHeight, _
        P.hDC, X - 9, Y - 7, 19, 15, vbSrcCopy
    frmMain.picScr.Line (sX * 9 - 1, Sy * 7 - 1)-(sX * 10, Sy * 8), vbBlack, B
    frmMain.picScr.Line (sX * 9 - 2, Sy * 7 - 2)-(sX * 10 + 1, Sy * 8 + 1), vbWhite, B
    frmMain.GetColorInfo frmMain.picScr.Point(sX * 9, Sy * 7)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteDC CBDC
    DeleteObject CBBitmap
End Sub
