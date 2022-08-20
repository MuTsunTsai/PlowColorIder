VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6075
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtResults 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "http://www.vb-helper.com"
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "URL"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STRING_SIZE = 128


' ------------------------------------------------------------------------------------------------------------
' Subroutine: GetWebData,ex. GetWebData("http://www.corel.com")
' Action:Return the source of a web page
' ------------------------------------------------------------------------------------------------------------
Private Function GetWebData(strUrl As String) As String
    Static hInternet, hSession, lngDataReturned As Long
    Static intReadFileResult As Integer
    Static strBuffer As String * 128
    Static strTotalData As String

    'retrieve a handle to the current internet session
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)

    If hSession = 0 Then
        MsgBox "Error opening Internet connection"
        Exit Function
    Else
        'retrieve a handle to the strUrl
        hInternet = InternetOpenUrl(hSession, strUrl, vbNullString, 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
    End If

    If hInternet = 0 Then
        MsgBox "Error opening Web page"
    Else
        'start reading the web page into a buffer, 128 bytes at a time
        intReadFileResult = InternetReadFile(hInternet, strBuffer, STRING_SIZE, lngDataReturned)
        
        'copy the contents of the buffer to strTotalData
        strTotalData = strBuffer
            
        'While there is still data left,
        Do While lngDataReturned <> 0
            'keep reading the web page into the buffer,
            intReadFileResult = InternetReadFile(hInternet, strBuffer, STRING_SIZE, lngDataReturned)
        
            'and keep appending the contents of strTotalData
            strTotalData = strTotalData + Mid(strBuffer, 1, lngDataReturned)
        Loop
    End If
   
    'return our internet handle
    intReadFileResult = InternetCloseHandle(hInternet)

    GetWebData = strTotalData

    'manually clear the string, just in case
    strTotalData = ""
End Function

Private Sub cmdLoad_Click()
    Screen.MousePointer = vbHourglass
    DoEvents

    txtResults.Text = GetWebData(txtURL.Text)

    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    Dim hgt As Single

    hgt = ScaleHeight - txtResults.Top
    If hgt < 120 Then hgt = 120
    txtResults.Move 0, txtResults.Top, ScaleWidth, hgt
End Sub
