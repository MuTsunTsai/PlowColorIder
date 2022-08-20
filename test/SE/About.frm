VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "關於"
   ClientHeight    =   2055
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4215
   ClipControls    =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1418.397
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   3958.103
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
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
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本"
      Height          =   225
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   2685
   End
   Begin VB.Label lblDescription 
      Caption         =   $"About.frx":0884
      ForeColor       =   &H00000000&
      Height          =   930
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
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub
