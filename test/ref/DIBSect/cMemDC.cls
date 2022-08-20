VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==================================================================================
' cDIBSection.cls
' Copyright © 1998 Steve Mcmahon
' Visit vbAccelerator at www.dogma.demon.co.uk
'
' Creates and manages a GDI Memory DC containing a
' compatible bitmap.
' ==================================================================================

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private m_hDC As Long
Private m_hBitmap As Long
Private m_hBitmapOld As Long
Private m_lWidth As Long
Private m_lHeight As Long

Public Function Create( _
      ByVal lHDC As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long _
   ) As Boolean
   Destroy
   m_hDC = CreateCompatibleDC(lHDC)
   If (m_hDC <> 0) Then
      m_lWidth = lWidth
      m_lHeight = lHeight
      m_hBitmap = CreateCompatibleBitmap(lHDC, m_lWidth, m_lHeight)
      If (m_hBitmap <> 0) Then
         m_hBitmapOld = SelectObject(m_hDC, m_hBitmap)
         Create = True
      Else
         DeleteDC m_hDC
         m_hDC = 0
      End If
   End If
End Function
Public Property Get Width() As Long
   Width = m_lWidth
End Property
Public Property Get Height() As Long
   Height = m_lHeight
End Property
Public Sub PaintPicture( _
        ByVal lHDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If (lDestWidth < 0) Then lDestWidth = m_lWidth
    If (lDestHeight < 0) Then lDestHeight = m_lHeight
    BitBlt lHDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
End Sub
Public Sub LoadPictureBlt( _
        ByVal lHDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If lSrcWidth < 0 Then lSrcWidth = m_lWidth
    If lSrcHeight < 0 Then lSrcHeight = m_lHeight
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lHDC, lSrcLeft, lSrcTop, eRop
End Sub
Public Sub Destroy()
   If (m_hDC <> 0) Then
      If (m_hBitmap <> 0) Then
         SelectObject m_hDC, m_hBitmapOld
         DeleteObject m_hBitmap
         m_hBitmap = 0
      End If
      DeleteDC m_hDC
   End If
End Sub

Private Sub Class_Terminate()
   Destroy
End Sub
