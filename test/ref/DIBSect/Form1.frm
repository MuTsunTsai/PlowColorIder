VERSION 5.00
Begin VB.Form frmDIBSection 
   Caption         =   "cDIBSection Tester"
   ClientHeight    =   7065
   ClientLeft      =   3630
   ClientTop       =   2010
   ClientWidth     =   7395
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   7395
   Begin VB.CheckBox chkUseMemDC 
      Caption         =   "Use Intermediate Memory DC"
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   4440
      Width           =   1275
   End
   Begin VB.OptionButton optSize 
      Caption         =   "512 x 512"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Width           =   1155
   End
   Begin VB.OptionButton optSize 
      Caption         =   "256 x 256"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fade in and ou&t"
      Height          =   675
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Emboss"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   5340
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Fade in and out with static"
      Height          =   675
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Resam&ple"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Load Audrey"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1500
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show &Static"
      Enabled         =   0   'False
      Height          =   375
      Left            =   300
      TabIndex        =   1
      Top             =   720
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Random DIB"
      Height          =   375
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "fps"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   1515
   End
End
Attribute VB_Name = "frmDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_cDIB As New cDIBSection
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub Command1_Click()
Dim lL As Long
    m_cDIB.Create 256, 256
    lL = (Command1.Left + Command1.Width) \ Screen.TwipsPerPixelX + 16
    m_cDIB.PaintPicture Me.hdc, lL
    m_cDIB.RandomiseBits 'True
    m_cDIB.PaintPicture Me.hdc, lL, m_cDIB.Height
    Command3.Enabled = True
    Command5.Enabled = True
End Sub

Private Sub Command2_Click()
Dim cDibDisp As cDIBSection
Dim cDibPic As cDIBSection
Dim cDC As cMemDC
Dim sPic As New StdPicture
Dim lAmount As Long
Dim lTIme As Long
Dim lFrames As Long
Dim lL As Long
Dim lDir As Long
Dim sFile As String
Dim bUseMemDC As Boolean
        
   lL = (Command1.Left + Command1.Width) \ Screen.TwipsPerPixelX + 16
    
   If (Command2.Caption <> "Fade in and ou&t") Then
      Command2.Caption = "Fade in and ou&t"
   Else
      Command2.Caption = "Stop"
      If (optSize(0).Value) Then
         sFile = App.Path & "\vbaccel.gif"
      Else
         sFile = App.Path & "\vbaccel2.jpg"
      End If
      Set sPic = LoadPicture(sFile)
      Set cDibPic = New cDIBSection
      Set cDibDisp = New cDIBSection
      cDibPic.CreateFromPicture sPic
      cDibDisp.Create cDibPic.Width, cDibPic.Height
      If (chkUseMemDC.Value = Checked) Then
         bUseMemDC = True
         Set cDC = New cMemDC
         cDC.Create Me.hdc, cDibPic.Width, cDibPic.Height
      End If
        
      lAmount = 255: lDir = -4
      Do While Command2.Caption = "Stop"
        
         If (timeGetTime - lTIme) > 1000 Then
             Label1.Caption = cDibPic.Width & "x" & cDibPic.Height & ": " & lFrames & "/second"
             lFrames = 0
             lTIme = timeGetTime
         Else
             lFrames = lFrames + 1
         End If
                  
         DoFade cDibPic, cDibDisp, lAmount
         
         If (bUseMemDC) Then
            cDC.LoadPictureBlt cDibDisp.hdc
            cDC.PaintPicture Me.hdc, lL
         Else
            cDibDisp.PaintPicture Me.hdc, lL
         End If

         lAmount = lAmount + lDir
         If (lAmount < 0) Then
            lDir = 4
            lAmount = 3
         ElseIf (lAmount = 255) Then
            lDir = -4
            lAmount = 255 - 4
         End If
         DoEvents
      Loop
      
   End If
   
End Sub

Private Sub Command3_Click()
Dim lL As Long
    lL = (Command1.Left + Command1.Width) \ Screen.TwipsPerPixelX + 16
If (Command3.Caption = "Show &Static") Then
    Command3.Caption = "Stop &Static"
Else
    Command3.Caption = "Show &Static"
End If
Do While Command3.Caption = "Stop &Static"
    m_cDIB.RandomiseBits True
    m_cDIB.PaintPicture Me.hdc, lL
    DoEvents
Loop
End Sub

Private Sub Command4_Click()
Dim sPic As StdPicture
Dim lL As Long
    lL = (Command1.Left + Command1.Width) \ Screen.TwipsPerPixelX + 16
    Set sPic = LoadPicture(App.Path & "\audrey.jpg")
    m_cDIB.CreateFromPicture sPic
    m_cDIB.PaintPicture Me.hdc, lL
    Command5.Enabled = True
    Command3.Enabled = True
End Sub

Private Sub Command5_Click()
Dim lL As Long
Dim cDib2 As cDIBSection
    lL = (Command1.Left + Command1.Width) \ Screen.TwipsPerPixelX + 16
    Set cDib2 = m_cDIB.Resample(m_cDIB.Height * 1.5, m_cDIB.Width * 1.5)
    Me.Cls
    m_cDIB.Create cDib2.Width, cDib2.Height
    cDib2.PaintPicture m_cDIB.hdc
    m_cDIB.PaintPicture Me.hdc, lL
    
End Sub

Private Sub Command6_Click()
Dim cDibPic As cDIBSection
Dim cDibDisp As cDIBSection
Dim cDC As cMemDC
Dim lAmount As Long, lAmount2 As Long
Dim lOffset As Long
Dim lRndAmount As Long
Dim lL As Long
Dim sPic As StdPicture
Dim lTIme As Long
Dim lFrames As Long
Dim sFile As String
Dim bUseMemDC As Boolean

    lL = (Command1.Left + Command1.Width) \ Screen.TwipsPerPixelX + 16
    
    If (Command6.Caption <> "&Fade in and out with static") Then
        Command6.Caption = "&Fade in and out with static"
    Else
      Command6.Caption = "Stop"
      If (optSize(0).Value) Then
         sFile = App.Path & "\vbaccel.gif"
      Else
         sFile = App.Path & "\vbaccel2.jpg"
      End If
      Set sPic = LoadPicture(sFile)
      Set cDibPic = New cDIBSection
      Set cDibDisp = New cDIBSection
      cDibPic.CreateFromPicture sPic
      cDibDisp.Create cDibPic.Width, cDibPic.Height
      
      If (chkUseMemDC.Value = Checked) Then
         bUseMemDC = True
         Set cDC = New cMemDC
         cDC.Create Me.hdc, cDibPic.Width, cDibPic.Height
      End If
    End If
    
    lAmount2 = 255
    Do While Command6.Caption = "Stop"
        If (timeGetTime - lTIme) > 1000 Then
            Label1.Caption = cDibPic.Width & "x" & cDibPic.Height & ": " & lFrames & "/second"
            lFrames = 0
            lTIme = timeGetTime
        Else
            lFrames = lFrames + 1
        End If
        If (lAmount < 251) Then
            lAmount = lAmount + 4
            DoStatic cDibPic, cDibDisp, lAmount, lOffset
        Else
            lAmount = 255
            If (lOffset < 251) Then
                lOffset = lOffset + 4
                DoStatic cDibPic, cDibDisp, lAmount, lOffset
            Else
                lOffset = 255
                If (lRndAmount < cDibPic.Height \ 6) Then
                    lRndAmount = lRndAmount + 1
                    BlowApart cDibPic, cDibDisp, lRndAmount
                Else
                    If (lAmount2 > 16) Then
                        lAmount2 = lAmount2 - 8
                        DoStatic cDibPic, cDibDisp, lAmount2, 0
                    Else
                        ' start again:
                        lAmount = 0: lOffset = 0: lRndAmount = 0: lAmount2 = 255
                        cDibPic.CreateFromPicture sPic
                    End If
                End If
            End If
        End If
        
      If (bUseMemDC) Then
         cDC.LoadPictureBlt cDibDisp.hdc
         cDC.PaintPicture Me.hdc, lL
      Else
         cDibDisp.PaintPicture Me.hdc, lL
      End If
      DoEvents
   Loop
End Sub

Private Sub Command7_Click()
Dim sPic As StdPicture
Dim cBuff As New cDIBSection
Dim cIP As New cImageProcessDIB
Dim lL As Long

    lL = (Command1.Left + Command1.Width) \ Screen.TwipsPerPixelX + 16

    Set sPic = LoadPicture(App.Path & "\vbaccel.gif")
    m_cDIB.CreateFromPicture sPic
    cBuff.Create m_cDIB.Width, m_cDIB.Height
    
    cIP.FilterType = eEmboss
    cIP.ProcessImage m_cDIB, cBuff
    m_cDIB.PaintPicture Me.hdc, lL
    
End Sub
