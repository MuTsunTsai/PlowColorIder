VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImageProcessDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Enum EFilterTypes
    eBlur
    eBlurMore
    eSoften
    eSoftenMore
    eSharpen
    eSharpenMore
    eUnSharp
    eEmboss
    eCustom
End Enum

Public Enum eFilterError
    eeFilterErrorBase = vbObjectError Or 1048 Or &H500
End Enum

Public Event InitProgress(ByVal lMax As Long)
Public Event Progress(ByVal lPosition As Long)
Public Event Complete(ByVal lTimeMs As Long)

Private m_iSize As Long
Private m_iOffset As Long
Private m_iFilt() As Long
Private m_iWeight As Long

Private m_eFilterType As EFilterTypes

Public Property Let FilterType(ByVal eType As EFilterTypes)
    m_eFilterType = eType
    If (m_eFilterType <> eCustom) Then
        pBuildFilterArray
    End If
End Property

Public Property Get FilterArraySize() As Long
    FilterArraySize = m_iSize
End Property
Public Property Let FilterArraySize(ByVal lSize As Long)
    If (lSize Mod 2) = 0 Then
        Err.Raise eeFilterErrorBase + 1, App.EXEName & ".cImageProcess", "Size must be an odd number"
    Else
        If (lSize < 0) Or (lSize > 9) Then
            Err.Raise eeFilterErrorBase + 2, App.EXEName & ".cImageProcess", "Invalid size.  Size should be an odd number from 3 to 9"
        Else
            m_iSize = lSize
            m_iOffset = m_iSize \ 2
            ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset) As Long
        End If
    End If
End Property
Public Property Get FilterValue(ByVal iX As Long, ByVal iY As Long) As Long
    FilterValue = m_iFilt(iX, iY)
End Property
Public Property Let FilterValue(ByVal iX As Long, ByVal iY As Long, ByVal lValue As Long)
    m_iFilt(iX, iY) = lValue
End Property
Public Property Get FilterWeight() As Long
    FilterWeight = m_iWeight
End Property
Public Property Let FilterWeight(lWeight As Long)
    m_iWeight = lWeight
End Property

Private Sub pBuildFilterArray()
Dim i As Long, j As Long
Dim iX As Long, iY As Long, iLM As Long

    m_iWeight = 0
    
    Select Case m_eFilterType
    Case eBlur, eBlurMore
        If (m_eFilterType = eBlur) Then
            FilterArraySize = 3
        Else
            FilterArraySize = 5
        End If
        For i = -m_iOffset To m_iOffset
            For j = -m_iOffset To m_iOffset
                m_iFilt(i, j) = 1
                m_iWeight = m_iWeight + m_iFilt(i, j)
            Next j
        Next i
        
    Case eSoften, eSoftenMore
        If (m_eFilterType = eSoften) Then
            FilterArraySize = 3
        Else
            FilterArraySize = 5
        End If
        For i = -m_iOffset To m_iOffset
            For j = -m_iOffset To m_iOffset
                
                iX = Abs(i)
                iY = Abs(j)
                If (iX > iY) Then
                    iLM = iX
                Else
                    iLM = iY
                End If
                
                If (iLM = 0) Then
                    m_iFilt(i, j) = (m_iSize * (m_iSize / 2#))
                Else
                    m_iFilt(i, j) = m_iOffset - iLM + 1
                End If
                Debug.Print m_iFilt(i, j); ",";
                m_iWeight = m_iWeight + m_iFilt(i, j)
            Next j
            Debug.Print
        Next i
        Debug.Print m_iWeight
        
    Case eSharpen, eSharpenMore
        FilterArraySize = 3
        If (m_eFilterType = eSharpen) Then
            m_iFilt(-1, -1) = -1: m_iFilt(-1, 0) = -1: m_iFilt(-1, 1) = -1
            m_iFilt(0, -1) = -1: m_iFilt(0, 0) = 15: m_iFilt(0, 1) = -1
            m_iFilt(1, -1) = -1: m_iFilt(1, 0) = -1: m_iFilt(1, 1) = -1
        Else
            m_iFilt(-1, -1) = 0: m_iFilt(-1, 0) = -1: m_iFilt(-1, 1) = 0
            m_iFilt(0, -1) = -1: m_iFilt(0, 0) = 5: m_iFilt(0, 1) = -1
            m_iFilt(1, -1) = 0: m_iFilt(1, 0) = -1: m_iFilt(1, 1) = 0
        End If
        For i = -m_iOffset To m_iOffset
            For j = -m_iOffset To m_iOffset
                m_iWeight = m_iWeight + m_iFilt(i, j)
            Next j
        Next i
        
    Case eEmboss
        FilterArraySize = 3
        m_iFilt(-1, -1) = -1: m_iFilt(1, 1) = 1
        m_iWeight = 1
        
    End Select
End Sub

Public Function ProcessImage( _
        ByRef cImage As cDIBSection, _
        ByRef cBuffer As cDIBSection _
    ) As Boolean
    Select Case m_eFilterType
    Case eBlur, eBlurMore, eCustom, eSharpen, eSharpenMore, eSoften, eSoftenMore
        ProcessImage = pbStandardFilter(cImage, cBuffer)
        cImage.LoadPictureBlt cBuffer.hDC

    Case eUnSharp
        ' Subtract a blurred version of the image from twice the
        ' original bitmap's value:
        FilterType = eBlur
        pbStandardFilter cImage, cBuffer
        AddImages cBuffer, cImage, -1, 0, 0, 0, 2, 0, 0, 0
        FilterType = eUnSharp
        
    Case eEmboss
        ' Perform emboss filter as normal, then add 127 to the R,G,B
        ' values to give a gray background
        ProcessImage = pbStandardFilter(cImage, cBuffer)
        AddImages cBuffer, cImage, 1, 127, 127, 127, 0, 0, 0, 0
    End Select
End Function
Private Function pbStandardFilter( _
        ByRef cImage As cDIBSection, _
        ByRef cBuffer As cDIBSection _
    ) As Boolean
' these are used to address the pixel using matrices
Dim pict() As Byte
Dim pict2() As Byte
Dim sa As SAFEARRAY2D
Dim sa2 As SAFEARRAY2D
Dim x As Long, y As Long
Dim rgbOffset As Long, xOffset As Long
Dim r As Long, g As Long, b As Long
Dim i As Long, j As Long, yMax As Long, xMax As Long
Dim lTIme As Long
Dim rR As Long, rB As Long, rG As Long
       
    ' have the local matrix point to bitmap pixels
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cImage.Height 'bmp.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cImage.BytesPerScanLine 'bmp.bmWidthBytes
        .pvData = cImage.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
        
    ' have the local matrix point to bitmap pixels
    With sa2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cBuffer.Height 'bmp2.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cBuffer.BytesPerScanLine
        .pvData = cBuffer.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
    
    ' Do filter on pict into pict2
    lTIme = timeGetTime
    
    rgbOffset = m_iOffset * 3
    yMax = UBound(pict, 2) - m_iOffset
    xMax = UBound(pict, 1) - 3 - rgbOffset
    RaiseEvent InitProgress(xMax)
    
    For x = rgbOffset To xMax Step 3
        For y = m_iOffset To yMax
            'Debug.Print X, Y
            'Debug.Print pict(X + i, Y + j), pict(X + 1 + i, Y + j), pict(X + 2 + i, Y + j)
            r = 0: g = 0: b = 0
            For i = -m_iOffset To m_iOffset
                xOffset = i * 3
                For j = -m_iOffset To m_iOffset
                    r = r + m_iFilt(i, j) * pict(x + xOffset, y + j)
                    g = g + m_iFilt(i, j) * pict(x + 1 + xOffset, y + j)
                    b = b + m_iFilt(i, j) * pict(x + 2 + xOffset, y + j)
                Next j
            Next i
            rR = r \ m_iWeight: rG = g \ m_iWeight: rB = b \ m_iWeight
                If (rR < 0) Then rR = 0
                If (rG < 0) Then rG = 0
                If (rB < 0) Then rB = 0
                If (rR > 255) Then rR = 255
                If (rG > 255) Then rG = 255
                If (rB > 255) Then rB = 255
            'Debug.Print rR, rG, rB, vbCrLf
            pict2(x, y) = rR: pict2(x + 1, y) = rG: pict2(x + 2, y) = rB
        Next y
        RaiseEvent Progress(x)
    Next x
         
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
        
    RaiseEvent Complete(timeGetTime - lTIme)
    
    pbStandardFilter = True
    
End Function
Public Function AddImages( _
        ByRef cFrom As cDIBSection, _
        ByRef cTo As cDIBSection, _
        ByVal lFromMultiplier As Long, _
        ByVal lFromOffsetR As Long, ByVal lFromOffsetG As Long, ByVal lFromOffsetB As Long, _
        ByVal lToMultiplier As Long, _
        ByVal lToOffsetR As Long, ByVal lToOffsetG As Long, ByVal lToOffsetB As Long _
    ) As Boolean
' these are used to address the pixel using matrices
Dim pict() As Byte
Dim pict2() As Byte
Dim sa As SAFEARRAY2D
Dim sa2 As SAFEARRAY2D
Dim x As Long, y As Long
Dim i As Long, j As Long, yMax As Long, lTIme As Long
Dim rR As Long, rG As Long, rB As Long

    lTIme = timeGetTime()
       
    ' have the local matrix point to bitmap pixels
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cTo.BytesPerScanLine
        .pvData = cTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    ' Pict now stores the To buffer
        
    ' have the local matrix point to bitmap pixels
    With sa2
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cFrom.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cFrom.BytesPerScanLine
        .pvData = cFrom.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4
    ' Pict2 now stores the From buffer

    yMax = cTo.Height - 1
    
    For x = 0 To (cTo.Width - 1) * 3 Step 3
        For y = 0 To yMax
            rR = (pict(x, y) + lToOffsetR) * lToMultiplier + (pict2(x, y) + lFromOffsetR) * lFromMultiplier
            rG = (pict(x + 1, y) + lToOffsetG) * lToMultiplier + (pict2(x + 1, y) + lFromOffsetG) * lFromMultiplier
            rB = (pict(x + 2, y) + lToOffsetB) * lToMultiplier + (pict2(x + 2, y) + lFromOffsetG) * lFromMultiplier
                If (rR < 0) Then rR = 0
                If (rG < 0) Then rG = 0
                If (rB < 0) Then rB = 0
                If (rR > 255) Then rR = 255
                If (rG > 255) Then rG = 255
                If (rB > 255) Then rB = 255
            
            pict(x, y) = rR
            pict(x + 1, y) = rG
            pict(x + 2, y) = rB
        Next y
        'prgMain.Value = x
    Next x
    
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
        
End Function



