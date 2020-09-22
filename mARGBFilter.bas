Attribute VB_Name = "mARGBFilter"
'================================================
' Module:        mARGBFilter.bas (reduced)
' Author:        Carles P.V.
' Dependencies:  cDIB.cls
' Last revision: 2004.06.14
'================================================
'
'   - Despeckle filter:
'     Based on article (http://www.dai.ed.ac.uk/HIPR2/crimmins.htm#2)

Option Explicit

'-- API:

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long



Public Sub Despeckle(oDIB32 As cDIB)
'-- Crimmins Speckle Removal

  Dim aBits() As RGBQUAD
  Dim uSA     As SAFEARRAY2D
  Dim x As Long, W As Long
  Dim y As Long, H As Long
  
  Dim pX As Byte             ' Current pixel
  Dim pA As Byte, pB As Byte ' Neighbors [N-S, W-S, NW-SE, SW-NE]
  
    If (oDIB32.hDIB) Then
        
        Call pvMapDIB(oDIB32, aBits(), uSA)
        
        W = oDIB32.Width - 2
        H = oDIB32.Height - 2
        
        '== N-S direction ==
        
        For y = 1 To H
            For x = 1 To W
                
                '-- Blue channel
                    pX = aBits(x, y).B
                    pA = aBits(x, y - 1).B
                    pB = aBits(x, y + 1).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).B = pX
                
                '-- Green channel
                    pX = aBits(x, y).G
                    pA = aBits(x, y - 1).G
                    pB = aBits(x, y + 1).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).G = pX
                    
                '-- Red channel
                    pX = aBits(x, y).R
                    pA = aBits(x, y - 1).R
                    pB = aBits(x, y + 1).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).R = pX
            Next x
        Next y
        
        '== W-E direction ==
        
        For y = 1 To H
            For x = 1 To W
                
                '-- Blue channel
                    pX = aBits(x, y).B
                    pA = aBits(x - 1, y).B
                    pB = aBits(x + 1, y).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).B = pX
                    
                '-- Green channel
                    pX = aBits(x, y).G
                    pA = aBits(x - 1, y).G
                    pB = aBits(x + 1, y).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).G = pX

                '-- Red channel
                    pX = aBits(x, y).R
                    pA = aBits(x - 1, y).R
                    pB = aBits(x + 1, y).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).R = pX
            Next x
        Next y
        
        '== NW-SE direction ==
        
        For y = 1 To H
            For x = 1 To W
                    
                '-- Blue channel:
                    pX = aBits(x, y).B
                    pA = aBits(x - 1, y - 1).B
                    pB = aBits(x + 1, y + 1).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).B = pX
                    
                '-- Green channel:
                    pX = aBits(x, y).G
                    pA = aBits(x - 1, y - 1).G
                    pB = aBits(x + 1, y + 1).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).G = pX
                    
                '-- Red channel:
                    pX = aBits(x, y).R
                    pA = aBits(x - 1, y - 1).R
                    pB = aBits(x + 1, y + 1).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).R = pX
            Next x
        Next y
        
        '== SW-NE direction ==
        
        For y = 1 To H
            For x = 1 To W
                    
                '-- Blue channel:
                    pX = aBits(x, y).B
                    pA = aBits(x - 1, y + 1).B
                    pB = aBits(x + 1, y - 1).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).B = pX
                    
                '-- Green channel:
                    pX = aBits(x, y).G
                    pA = aBits(x - 1, y + 1).G
                    pB = aBits(x + 1, y - 1).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).G = pX

                '-- Red channel:
                    pX = aBits(x, y).R
                    pA = aBits(x - 1, y + 1).R
                    pB = aBits(x + 1, y - 1).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    aBits(x, y).R = pX
            Next x
        Next y

        Call pvUnmapDIB(aBits())
    End If
End Sub

Public Sub Sharpen(oDIB32 As cDIB, Optional ByVal Level As Long = 75)

  Dim sDIB32   As New cDIB
  
  Dim asBits() As RGBQUAD
  Dim adBits() As RGBQUAD
  Dim suSA     As SAFEARRAY2D
  Dim duSA     As SAFEARRAY2D
  
  Dim x As Long, W As Long
  Dim y As Long, H As Long
  Dim i As Long, j As Long
  Dim rctImg As RECT2
  
  Dim R As Long
  Dim G As Long
  Dim B As Long
  
  Dim Lev  As Long
  Dim Wgt  As Long
  
    If (oDIB32.hDIB) Then
        
        Call oDIB32.CloneTo(sDIB32)
        
        Call pvMapDIB(sDIB32, asBits(), suSA)
        Call pvMapDIB(oDIB32, adBits(), duSA)
            
        Call SetRect(rctImg, 0, 0, oDIB32.Width, oDIB32.Height)
        W = oDIB32.Width - 1
        H = oDIB32.Height - 1
    
        Lev = 109 - Level
        Wgt = Lev - 8
        
        For y = 0 To H
            For x = 0 To W
                
                Wgt = Lev
                B = Lev * CLng(asBits(x, y).B)
                G = Lev * CLng(asBits(x, y).G)
                R = Lev * CLng(asBits(x, y).R)
                
                For j = -1 To 1
                    For i = -1 To 1
                    
                        If (PtInRect(rctImg, x + i, y + j)) Then
                            With asBits(x + i, y + j)
                                If (.A) Then
                                    R = R - .R
                                    G = G - .G
                                    B = B - .B
                                    Wgt = Wgt - 1
                                End If
                            End With
                        End If
                    Next i
                Next j
                
                B = B / Wgt
                G = G / Wgt
                R = R / Wgt
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                
                With adBits(x, y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next x
        Next y
        Call pvUnmapDIB(asBits())
        Call pvUnmapDIB(adBits())
    End If
End Sub

Public Sub Soften(oDIB32 As cDIB, Optional ByVal Level As Long = 75)

  Dim sDIB32  As New cDIB
  
  Dim asBits() As RGBQUAD
  Dim adBits() As RGBQUAD
  Dim suSA     As SAFEARRAY2D
  Dim duSA     As SAFEARRAY2D
  
  Dim x As Long, W As Long
  Dim y As Long, H As Long
  Dim i As Long, j As Long
  Dim rctImg As RECT2
  
  Dim R As Long
  Dim G As Long
  Dim B As Long
  
  Dim Lev As Long
  Dim Wgt As Long
  
    If (oDIB32.hDIB) Then
        
        Call oDIB32.CloneTo(sDIB32)
        
        Call pvMapDIB(sDIB32, asBits(), suSA)
        Call pvMapDIB(oDIB32, adBits(), duSA)
        
        Call SetRect(rctImg, 0, 0, oDIB32.Width, oDIB32.Height)
        W = oDIB32.Width - 1
        H = oDIB32.Height - 1
    
        Lev = 101 - Level
        Wgt = Lev + 4
        
        For y = 0 To H
            For x = 0 To W
                
                Wgt = Lev
                B = Lev * CLng(asBits(x, y).B)
                G = Lev * CLng(asBits(x, y).G)
                R = Lev * CLng(asBits(x, y).R)
                
                For j = -1 To 1
                    For i = -1 To 1
                    
                        If (i <> j And PtInRect(rctImg, x + i, y + j)) Then
                            With asBits(x + i, y + j)
                                If (.A) Then
                                    R = R + .R
                                    G = G + .G
                                    B = B + .B
                                    Wgt = Wgt + 1
                                End If
                            End With
                        End If
                    Next i
                Next j
                
                B = B / Wgt
                G = G / Wgt
                R = R / Wgt
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                
                With adBits(x, y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next x
        Next y
        Call pvUnmapDIB(asBits())
        Call pvUnmapDIB(adBits())
    End If
End Sub

Public Sub Colorize(oDIB32 As cDIB, ByVal Color As Long)

  Dim aBits() As RGBQUAD
  Dim uSA     As SAFEARRAY2D
  Dim x As Long, W As Long
  Dim y As Long, H As Long
  
  Dim lH As Long
  Dim lS As Long
  Dim lL As Long

    If (oDIB32.hDIB) Then

        Call pvMapDIB(oDIB32, aBits(), uSA)
        Call pvRGBToHSL(Color And &HFF&, (Color And &HFF00&) \ 256, (Color And &HFF0000) \ 65536, lH, lS, lL)
        
        W = oDIB32.Width - 1
        H = oDIB32.Height - 1
        
        For y = 0 To H
            For x = 0 To W
                
                With aBits(x, y)
                    lL = (299& * .R + 587& * .G + 114& * .B) \ 2550
                    Call pvHSLToRGB(lH, lS, lL, .R, .G, .B)
                End With
            Next x
        Next y
    End If
End Sub

Public Sub Greys(oDIB32 As cDIB)

  Dim aBits() As RGBQUAD
  Dim uSA     As SAFEARRAY2D
  Dim x As Long, W As Long
  Dim y As Long, H As Long
  
  Dim L As Byte

    If (oDIB32.hDIB) Then

        Call pvMapDIB(oDIB32, aBits(), uSA)
        
        W = oDIB32.Width - 1
        H = oDIB32.Height - 1
        
        For y = 0 To H
            For x = 0 To W
                
                With aBits(x, y)
                    L = (299& * .R + 587& * .G + 114& * .B) \ 1000
                    .R = L
                    .G = L
                    .B = L
                End With
            Next x
        Next y
    End If
End Sub

Public Sub DropShadow(oDIB32 As cDIB, _
                      Optional ByVal xOffset As Long = 1, _
                      Optional ByVal yOffset As Long = 1, _
                      Optional ByVal Opacity As Long = 75)

  Dim oDIBShadow    As New cDIB
  Dim oDIBBuffer    As New cDIB
  Dim uSABits       As SAFEARRAY2D
  Dim aBits()       As RGBQUAD
  Dim uSABitsShadow As SAFEARRAY2D
  Dim aBitsShadow() As RGBQUAD
  
  Dim hBitmap       As Long
  Dim hGraphics     As Long
  Dim hBitmapShadow As Long
  Dim hBitmapBuffer As Long
  Dim uMatrix       As COLORMATRIX
  Dim hAttributes   As Long
  
  Dim x As Long, W As Long
  Dim y As Long, H As Long
  Dim i As Long, j As Long
  
  Dim rctImg As RECT2
  Dim lAcm   As Long
  Dim lPts   As Long
    
    If (oDIB32.hDIB) Then
    
        With oDIB32
        
            Call .CloneTo(oDIBBuffer)
            Call oDIBShadow.Create(.Width, .Height, [32_bpp])
            Call oDIBShadow.Reset
            
            Call pvMapDIB(oDIB32, aBits(), uSABits)
            Call pvMapDIB(oDIBShadow, aBitsShadow(), uSABitsShadow)
            
            Call SetRect(rctImg, 0, 0, .Width, .Height)
            W = .Width - 1
            H = .Height - 1
                
            For y = 0 To H
                For x = 0 To W
                
                    lPts = 1
                    lAcm = aBits(x, y).A
                
                    For j = -1 To 1
                        For i = -1 To 1
                            
                            If (PtInRect(rctImg, x + i, y + j)) Then
                                lAcm = lAcm + aBits(x + i, y + j).A
                                lPts = lPts + 1
                            End If
                        Next i
                    Next j
                    aBitsShadow(x, y).A = lAcm \ lPts
                Next x
            Next y
            
            Call pvUnmapDIB(aBits())
            Call pvUnmapDIB(aBitsShadow())
            
            Call .Reset
            
            Call ARGBBitmapFromGDIDIB32(oDIB32, hBitmap)
            Call ARGBBitmapFromGDIDIB32(oDIBShadow, hBitmapShadow)
            Call ARGBBitmapFromGDIDIB32(oDIBBuffer, hBitmapBuffer)
            Call GdipGetImageGraphicsContext(hBitmap, hGraphics)
                
            With uMatrix
                .m(0, 0) = 1
                .m(1, 1) = 1
                .m(2, 2) = 1
                .m(3, 3) = Opacity / 100
                .m(4, 4) = 1
            End With
            Call GdipCreateImageAttributes(hAttributes)
            Call GdipSetImageAttributesColorMatrix(hAttributes, [ColorAdjustTypeDefault], True, uMatrix, ByVal 0, [ColorMatrixFlagsDefault])
        
            Call GdipDrawImageRectRectI(hGraphics, hBitmapShadow, xOffset, yOffset, .Width, .Height, 0, 0, .Width, .Height, [UnitPixel], hAttributes, 0, 0)
            Call GdipDrawImageRectI(hGraphics, hBitmapBuffer, 0, 0, .Width, .Height)
            
            Call ARGBBitmapToGDIDIB32(oDIB32, hBitmap)
            
            Call GdipDisposeImage(hBitmap)
            Call GdipDisposeImage(hBitmapShadow)
            Call GdipDisposeImage(hBitmapBuffer)
            Call GdipDeleteGraphics(hGraphics)
            Call GdipDisposeImageAttributes(hAttributes)
        End With
    End If
End Sub

Public Sub FadeDownAlpha(oDIB32 As cDIB)

  Dim x As Long, W As Long
  Dim y As Long, H As Long
  
  Dim uSABits As SAFEARRAY2D
  Dim aBits() As RGBQUAD
  
    If (oDIB32.hDIB) Then
    
        Call pvMapDIB(oDIB32, aBits(), uSABits)
        
        W = oDIB32.Width - 1
        H = oDIB32.Height - 1
        
        For y = 0 To H
            For x = 0 To W
                aBits(x, y).A = (CLng(aBits(x, y).A) * y) \ H
            Next x
        Next y
        
        Call pvUnmapDIB(aBits())
    End If
End Sub

Public Sub AlphaScanlines(oDIB32 As cDIB)

  Dim x As Long, W As Long
  Dim y As Long, H As Long
  
  Dim uSABits As SAFEARRAY2D
  Dim aBits() As RGBQUAD
  
    If (oDIB32.hDIB) Then
    
        Call pvMapDIB(oDIB32, aBits(), uSABits)
        
        W = oDIB32.Width - 1
        H = oDIB32.Height - 1
        
        For y = 0 To H Step 2
            For x = 0 To W
                aBits(x, y).A = aBits(x, y).A \ 2
            Next x
        Next y
        
        Call pvUnmapDIB(aBits())
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvHSLToRGB(ByVal H As Long, ByVal S As Long, ByVal L As Long, R As Byte, G As Byte, B As Byte)
' by Donald (Sterex 1996), donald@xbeat.net, 20011124
  
  Dim lMax As Long, lMid As Long, lMin As Long
  Dim q As Single

    lMax = (L * 255) / 100
    
    If (S > 0) Then
      
        lMin = (100 - S) * lMax / 100
        q = (lMax - lMin) / 60
        
        Select Case H
        Case 0 To 60
            lMid = (H - 0) * q + lMin
            R = lMax: G = lMid: B = lMin
        Case 60 To 120
            lMid = -(H - 120) * q + lMin
            R = lMid: G = lMax: B = lMin
        Case 120 To 180
            lMid = (H - 120) * q + lMin
            R = lMin: G = lMax: B = lMid
        Case 180 To 240
            lMid = -(H - 240) * q + lMin
            R = lMin: G = lMid: B = lMax
        Case 240 To 300
            lMid = (H - 240) * q + lMin
            R = lMid: G = lMin: B = lMax
        Case 300 To 360
            lMid = -(H - 360) * q + lMin
            R = lMax: G = lMin: B = lMid
        End Select
    
      Else
        
        R = lMax: G = lMax: B = lMax
    End If
End Sub

Private Sub pvRGBToHSL(ByVal R As Long, ByVal G As Long, ByVal B As Long, H As Long, S As Long, L As Long)
' by Paul - wpsjr1@syix.com, 20011120
  
  Dim lMax         As Long
  Dim lMin         As Long
  Dim q            As Single
  Dim lDifference  As Long
  Static Lum(255)  As Long
  Static QTab(255) As Single
  Static Init      As Long
  
    If (Init = 0) Then
        For Init = 2 To 255 ' 0 and 1 are both 0
            Lum(Init) = Init * 100 / 255
        Next Init
        For Init = 1 To 255
            QTab(Init) = 60 / Init
        Next Init
    End If
    
    If (R > G) Then
        lMax = R: lMin = G
      Else
        lMax = G: lMin = R
    End If
    If (B > lMax) Then
        lMax = B
      ElseIf B < lMin Then
        lMin = B
    End If
    
    L = Lum(lMax)

    lDifference = lMax - lMin
    If (lDifference) Then
        
        S = (lDifference) * 100 / lMax ' do a 65K 2D lookup table here for more speed if needed
        
        q = QTab(lDifference)
        Select Case lMax
            Case R
                If (B > G) Then
                    H = q * (G - B) + 360
                  Else
                    H = q * (G - B)
                End If
            Case G
                H = q * (B - R) + 120
            Case B
                H = q * (R - G) + 240
        End Select
    End If
End Sub

Private Sub pvMapDIB(oDIB32 As cDIB, aBits() As RGBQUAD, uSA As SAFEARRAY2D)
    
    With uSA
        .cbElements = IIf(App.LogMode = 1, 1, 4)
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB32.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB32.Width
        .pvData = oDIB32.lpBits
    End With
    Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(uSA), 4)
End Sub

Private Sub pvUnmapDIB(aBits() As RGBQUAD)

    Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
End Sub
