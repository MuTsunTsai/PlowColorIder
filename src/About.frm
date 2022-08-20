VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "關於"
   ClientHeight    =   2295
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4350
   ClipControls    =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   153
   ScaleMode       =   3  '像素
   ScaleWidth      =   290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin 北斗色彩識別器.Weblink Weblink1 
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Location        =   "mailto:stargazer@abstreamace.com"
      Caption         =   "stargazer@abstreamace.com"
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  '沒有框線
      ClipControls    =   0   'False
      Height          =   480
      Left            =   120
      Picture         =   "About.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  '使用者自訂
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "關閉(&C)"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   945
   End
   Begin 北斗色彩識別器.Weblink Weblink2 
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Location        =   "http://www.abstreamace.com/sglab/"
      Caption         =   "http://www.abstreamace.com/sglab/"
   End
   Begin VB.Label Label1 
      Caption         =   "亦可至官方網站了解更多訊息，網址位於："
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本"
      Height          =   225
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   2685
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   3285
   End
   Begin VB.Label lblTitle 
      Caption         =   "名稱"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   2685
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "關於 " & App.Title
    lblDescription.Caption = "本程式由星君撰寫，為免費共享軟體，" & vbCrLf & _
        "敬請多加散播分享。祝您使用愉快。" & vbCrLf & "歡迎來信指教："
    lblVersion.Caption = "版本 " & getAppVersion(True)
    lblTitle.Caption = App.Title
End Sub
