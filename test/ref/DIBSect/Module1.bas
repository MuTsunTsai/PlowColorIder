Attribute VB_Name = "mDIBSectEffects"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Public Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long


Public Sub BlowApart(ByRef cDibPic As cDIBSection, ByRef cDibDisp As cDIBSection, ByVal lAmount As Long)
Dim tSAPic As SAFEARRAY2D
Dim tSADisp As SAFEARRAY2D
Dim bPic() As Byte
Dim bDisp() As Byte
Dim x As Long, y As Long
Dim xC As Long, yC As Long
Dim xNew As Long, yNew As Long
Dim xEnd As Long, yEnd As Long
Dim bFinish As Boolean
    
    ' Get the bits in the from DIB section:
    With tSAPic
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibPic.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibPic.BytesPerScanLine()
        .pvData = cDibPic.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bPic()), VarPtr(tSAPic), 4

    With tSADisp
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibDisp.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibDisp.BytesPerScanLine()
        .pvData = cDibDisp.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDisp()), VarPtr(tSADisp), 4

    ' Copy the display picture to the dib:
    cDibPic.LoadPictureBlt cDibDisp.hDC

    xEnd = (cDibPic.Width - 1) * 3
    yEnd = cDibPic.Height - 1
    xC = xEnd \ 2
    yC = yEnd \ 2
    
    For y = 0 To yEnd
        For x = 0 To xEnd Step 3
            If (bPic(x, y) <> 0) Then
                bFinish = False
                If (x > xC) Then
                    xNew = x + Rnd * lAmount * 3
                    If (xNew > xEnd) Then
                        bFinish = True
                    End If
                Else
                    xNew = x - Rnd * lAmount * 3
                    If (xNew < 0) Then
                        bFinish = True
                    End If
                End If
                
                If (y < yC) Then
                    yNew = y - Rnd * lAmount
                    If (yNew < 0) Then
                        bFinish = True
                    End If
                Else
                    yNew = y + Rnd * lAmount
                    If (yNew > yEnd) Then
                        bFinish = True
                    End If
                End If
                
                If Not (bFinish) Then
                    bDisp(xNew, yNew) = bPic(x, y)
                    bDisp(xNew + 1, yNew) = bPic(x + 1, y)
                    bDisp(xNew + 2, yNew) = bPic(x + 2, y)
                    
                    bPic(xNew, yNew) = bPic(x, y)
                    bPic(xNew + 1, yNew) = bPic(x + 1, y)
                    bPic(xNew + 2, yNew) = bPic(x + 2, y)
                End If
            End If
        Next x
    Next y
    
    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bPic), 0&, 4
    CopyMemory ByVal VarPtrArray(bDisp), 0&, 4


End Sub
Public Sub DoStatic(ByRef cDibPic As cDIBSection, ByRef cDibDisp As cDIBSection, ByVal lAmount As Long, ByVal lOffset As Long)
Dim tSAPic As SAFEARRAY2D
Dim tSADisp As SAFEARRAY2D
Dim bPic() As Byte
Dim bDisp() As Byte
Dim x As Long, y As Long
Dim lRnd As Long
Dim xEnd As Long
Dim lR As Long, lG As Long, lB As Long
    
    ' Get the bits in the from DIB section:
    With tSAPic
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibPic.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibPic.BytesPerScanLine()
        .pvData = cDibPic.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bPic()), VarPtr(tSAPic), 4

    With tSADisp
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibDisp.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibDisp.BytesPerScanLine()
        .pvData = cDibDisp.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDisp()), VarPtr(tSADisp), 4

    xEnd = (cDibPic.Width - 1) * 3
    For y = 0 To cDibPic.Height - 1
        For x = 0 To xEnd Step 3
            'If (bPic(x, y) <> 0) Or (bPic(x + 1, y) <> 0) Or (bPic(x + 2, y) <> 0) Then
                lRnd = Rnd * (lAmount - lOffset)
                lB = (lRnd + lOffset) * bPic(x, y) \ 255
                lG = (lRnd + lOffset) * bPic(x + 1, y) \ 255
                lR = (lRnd + lOffset) * bPic(x + 2, y) \ 255
                bDisp(x, y) = lB
                bDisp(x + 1, y) = lG
                bDisp(x + 2, y) = lR
            'End If
        Next x
    Next y

    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bPic), 0&, 4
    CopyMemory ByVal VarPtrArray(bDisp), 0&, 4

End Sub
Public Sub DoFade(ByRef cDibPic As cDIBSection, ByRef cDibDisp As cDIBSection, ByVal lAmount As Long)
Dim tSAPic As SAFEARRAY2D
Dim tSADisp As SAFEARRAY2D
Dim bPic() As Byte
Dim bDisp() As Byte
Dim x As Long, y As Long
Dim lRnd As Long
Dim xEnd As Long
Dim lR As Long, lG As Long, lB As Long
    
    ' Get the bits in the from DIB section:
    With tSAPic
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibPic.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibPic.BytesPerScanLine()
        .pvData = cDibPic.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bPic()), VarPtr(tSAPic), 4

    With tSADisp
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibDisp.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibDisp.BytesPerScanLine()
        .pvData = cDibDisp.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDisp()), VarPtr(tSADisp), 4

    xEnd = (cDibPic.Width - 1) * 3
    For y = 0 To cDibPic.Height - 1
        For x = 0 To xEnd Step 3
            lB = lAmount * bPic(x, y) \ 255
            lG = lAmount * bPic(x + 1, y) \ 255
            lR = lAmount * bPic(x + 2, y) \ 255
            bDisp(x, y) = lB
            bDisp(x + 1, y) = lG
            bDisp(x + 2, y) = lR
        Next x
    Next y

    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bPic), 0&, 4
    CopyMemory ByVal VarPtrArray(bDisp), 0&, 4

End Sub


