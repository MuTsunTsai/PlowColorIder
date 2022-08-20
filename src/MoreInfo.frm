VERSION 5.00
Begin VB.Form MoreInfo 
   BorderStyle     =   1  '單線固定
   Caption         =   "更多資訊"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MoreInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  '像素
   ScaleWidth      =   257
   StartUpPosition =   2  '螢幕中央
   Begin 北斗色彩識別器.PlowFlatButton L 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "按一下以複製到剪貼簿"
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "三原色百分比：RGB()"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "關閉(&C)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2520
      Width           =   945
   End
   Begin 北斗色彩識別器.PlowFlatButton L 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "按一下以複製到剪貼簿"
      Top             =   375
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "HSV 值："
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin 北斗色彩識別器.PlowFlatButton L 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "按一下以複製到剪貼簿"
      Top             =   630
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "HSL 值："
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin 北斗色彩識別器.PlowFlatButton L 
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "按一下以複製到剪貼簿"
      Top             =   885
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "YUV 值："
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin 北斗色彩識別器.PlowFlatButton L 
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "按一下以複製到剪貼簿"
      Top             =   1140
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "CMYK 值："
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin 北斗色彩識別器.PlowFlatButton L 
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "按一下以複製到剪貼簿"
      Top             =   1650
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "VC 色碼："
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin 北斗色彩識別器.PlowFlatButton L 
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "按一下以複製到剪貼簿"
      Top             =   1905
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "VB 色碼："
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin 北斗色彩識別器.PlowFlatButton L 
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "按一下以複製到剪貼簿"
      Top             =   2160
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "PASCAL 色碼："
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin 北斗色彩識別器.PlowFlatButton L 
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "按一下以複製到剪貼簿（本數值受模擬環境參數影響，計算結果僅供參考）"
      Top             =   1395
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "CIE Lab 值："
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MoreInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Static C As New ColorInfo, D As New ColorHSV, E As New ColorHSL
    Static F As New ColorCMYK, G As New ColorYUV, H As New ColorLab
    C.Color = frmMain.TColor.BackColor: D.Color = frmMain.TColor.BackColor
    E.Color = frmMain.TColor.BackColor: F.Color = frmMain.TColor.BackColor
    G.Color = frmMain.TColor.BackColor: H.Color = frmMain.TColor.BackColor
    L(0).Tag = "RGB(" & C.getRPer & ", " & C.getGPer & ", " & C.getBPer & ")"
    L(0).Caption = "三原色百分比︰" & L(0).Tag
    L(1).Tag = "HSV(" & Mid(D.H, 1, 5) & ", " & Mid(D.S, 1, 5) & "%, " & Mid(D.I, 1, 5) & "%)"
    L(1).Caption = "HSV 值：" & L(1).Tag
    L(2).Tag = "HSL(" & Mid(E.H, 1, 5) & ", " & Mid(E.S, 1, 5) & "%, " & Mid(E.I, 1, 5) & "%)"
    L(2).Caption = "HSL 值：" & L(2).Tag
    L(3).Tag = Mid(G.Y, 1, 5) & "%, " & Mid(G.U, 1, 5) & "%, " & Mid(G.V, 1, 5) & "%"
    L(3).Caption = "YUV 值：" & L(3).Tag
    L(4).Tag = Mid(F.C, 1, 5) & "%, " & Mid(F.m, 1, 5) & "%, " & Mid(F.Y, 1, 5) & "%, " & Mid(F.K, 1, 5) & "%"
    L(4).Caption = "CMYK 值：" & L(4).Tag
    L(5).Tag = "0x00" & C.getBBGGRR
    L(5).Caption = "VC 色碼：" & L(5).Tag
    L(6).Tag = "&H00" & C.getBBGGRR & "&"
    L(6).Caption = "VB 色碼：" & L(6).Tag
    L(7).Tag = "$00" & C.getBBGGRR
    L(7).Caption = "PASCAL 色碼：" & L(7).Tag
    L(8).Tag = Mid(H.CIEL, 1, 5) & ", " & Mid(H.CIEa, 1, 5) & ", " & Mid(H.CIEb, 1, 5)
    L(8).Caption = "CIE Lab 值：" & L(8).Tag
End Sub

Private Sub L_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then L(Index).ForeColor = vbBlue
End Sub

Private Sub L_MouseOut(Index As Integer)
    L(Index).ForeColor = vbBlack
End Sub

Private Sub L_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static txt As String, C As New ColorInfo
    If Button = 1 Then
        Clipboard.Clear
        Clipboard.SetText L(Index).Tag
        L_MouseOut Index
    End If
End Sub
