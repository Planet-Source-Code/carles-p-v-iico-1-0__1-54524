VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBDither"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'================================================
' Class:         cDIBDither.cls
' Author:        Carles P.V.
' Dependencies:  cDIB.cls
'                cPalette.cls
' Last revision: -
'================================================

Option Explicit

'-- API:

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

'//

'-- Private Variables:
Private m_Pow2(31)  As Long



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Initialize()

  Dim lIdx As Long

    For lIdx = 0 To 30
        m_Pow2(lIdx) = 2 ^ lIdx
    Next lIdx
    m_Pow2(31) = &H80000000
End Sub

Private Sub Class_Terminate()
    Erase m_Pow2()
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub DitherToGreyScale(oDIB32In As cDIB, oDIBOut As cDIB, Optional ByVal Diffuse As Boolean = False)
'-- Floyd-Steinberg error diffusion

  Dim aBitsIn()  As Byte
  Dim uSAIn      As SAFEARRAY2D
  Dim aBitsOut() As Byte
  Dim uSAOut     As SAFEARRAY2D

  Dim x As Long, xIn As Long
  Dim y As Long, yIn As Long
  Dim W As Long
  Dim H As Long

  Dim L As Byte

  Dim errR As Long, errG As Long, errB As Long
  Dim newR As Long, newG As Long, newB As Long

    If (oDIB32In.BPP = [32_bpp]) Then

        Call pvMapDIB(oDIB32In, aBitsIn(), uSAIn)
        Call pvMapDIB(oDIBOut, aBitsOut(), uSAOut)

        W = oDIB32In.Width - 1
        H = oDIB32In.Height - 1

        For y = 0 To H
            For x = 0 To W
            
                '-- Get grey level
                xIn = 4 * x
                L = (299 * CLng(aBitsIn(xIn + 2, y)) + _
                     587 * CLng(aBitsIn(xIn + 1, y)) + _
                     114 * CLng(aBitsIn(xIn + 0, y)) _
                     ) \ 1000
                
                '-- Pre-dither source
                aBitsIn(xIn + 0, y) = L
                aBitsIn(xIn + 1, y) = L
                aBitsIn(xIn + 2, y) = L
                
                '-- Set target index/color
                Select Case oDIBOut.BPP
                
                    Case [01_bpp]
                        
                        xIn = x \ 8
                        If (L < 128) Then
                            aBitsOut(xIn, y) = aBitsOut(xIn, y) And Not m_Pow2(7 - (x Mod 8))
                            L = 0
                          Else
                            aBitsOut(xIn, y) = aBitsOut(xIn, y) Or m_Pow2(7 - (x Mod 8))
                            L = 255
                        End If
                    
                    Case [04_bpp]
                        
                        xIn = x \ 2
                        If (x Mod 2 = 0) Then
                            aBitsOut(xIn, y) = (aBitsOut(xIn, y) And &HF) Or L * &H10
                          Else
                            aBitsOut(xIn, y) = (aBitsOut(xIn, y) And &HF0) Or L
                        End If
                        L = L * 17

                    Case [08_bpp]
                    
                        aBitsOut(x, y) = L

                    Case [24_bpp]
                    
                        xIn = 3 * x
                        aBitsOut(xIn + 0, y) = L
                        aBitsOut(xIn + 1, y) = L
                        aBitsOut(xIn + 2, y) = L

                    Case [32_bpp]
                    
                        xIn = 4 * x
                        aBitsOut(xIn + 0, y) = L
                        aBitsOut(xIn + 1, y) = L
                        aBitsOut(xIn + 2, y) = L
                End Select
                        
                '-- Diffuse error
                If (Diffuse) Then
                    
                    xIn = 4 * x
                    errB = CLng(aBitsIn(xIn + 0, y)) - L
                    errG = CLng(aBitsIn(xIn + 1, y)) - L
                    errR = CLng(aBitsIn(xIn + 2, y)) - L
                    aBitsIn(xIn + 0, y) = L
                    aBitsIn(xIn + 1, y) = L
                    aBitsIn(xIn + 2, y) = L

                    '-- Floyd-Steinberg error diffusion...
                    If (Abs(errB) + Abs(errG) + Abs(errR) > 3) Then
                        If (x < W) Then
                            xIn = 4 * x + 4
                            newB = aBitsIn(xIn + 0, y) + (7 * errB) / 16
                            newG = aBitsIn(xIn + 1, y) + (7 * errG) / 16
                            newR = aBitsIn(xIn + 2, y) + (7 * errR) / 16
                            If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                            If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                            If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                            aBitsIn(xIn + 0, y) = newB
                            aBitsIn(xIn + 1, y) = newG
                            aBitsIn(xIn + 2, y) = newR
                        End If
                        If (y < H) Then
                            xIn = 4 * x
                            yIn = y + 1
                            newB = aBitsIn(xIn + 0, yIn) + (5 * errB) / 16
                            newG = aBitsIn(xIn + 1, yIn) + (5 * errG) / 16
                            newR = aBitsIn(xIn + 2, yIn) + (5 * errR) / 16
                            If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                            If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                            If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                            aBitsIn(xIn + 0, yIn) = newB
                            aBitsIn(xIn + 1, yIn) = newG
                            aBitsIn(xIn + 2, yIn) = newR
                            If (x < W) Then
                                xIn = 4 * x + 4
                                newB = aBitsIn(xIn + 0, yIn) + errB / 16
                                newG = aBitsIn(xIn + 1, yIn) + errG / 16
                                newR = aBitsIn(xIn + 2, yIn) + errR / 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                aBitsIn(xIn + 0, yIn) = newB
                                aBitsIn(xIn + 1, yIn) = newG
                                aBitsIn(xIn + 2, yIn) = newR
                            End If
                            If (x > 0) Then
                                xIn = 4 * x - 4
                                newB = aBitsIn(xIn + 0, yIn) + (3 * errB) / 16
                                newG = aBitsIn(xIn + 1, yIn) + (3 * errG) / 16
                                newR = aBitsIn(xIn + 2, yIn) + (3 * errR) / 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                aBitsIn(xIn + 0, yIn) = newB
                                aBitsIn(xIn + 1, yIn) = newG
                                aBitsIn(xIn + 2, yIn) = newR
                            End If
                        End If
                    End If
                  Else
                    aBitsIn(xIn + 0, y) = L
                    aBitsIn(xIn + 1, y) = L
                    aBitsIn(xIn + 2, y) = L
                End If
            Next x
        Next y
        Call pvUnmapDIB(aBitsIn())
        Call pvUnmapDIB(aBitsOut())
    End If
End Sub

Public Sub DitherToColorPalette(oPalIn As cPalette, oDIB32In As cDIB, oDIBOut As cDIB, Optional ByVal Diffuse As Boolean = False)
'-- Floyd-Steinberg error diffusion

  Dim aBitsIn()  As Byte
  Dim uSAIn      As SAFEARRAY2D
  Dim aBitsOut() As Byte
  Dim uSAOut     As SAFEARRAY2D

  Dim x As Long, xIn As Long
  Dim y As Long, yIn As Long
  Dim W As Long
  Dim H As Long

  Dim aPalette() As Byte
  Dim aEntry     As Byte
  Dim lEntry     As Long

  Dim errR As Long, errG As Long, errB As Long
  Dim newR As Long, newG As Long, newB As Long

    If (oDIB32In.BPP = [32_bpp]) Then

        Call pvMapDIB(oDIB32In, aBitsIn(), uSAIn)
        Call pvMapDIB(oDIBOut, aBitsOut(), uSAOut)

        W = oDIB32In.Width - 1
        H = oDIB32In.Height - 1

        '-- Speed up for GetNearestPaletteIndex
        ReDim aPalette(4 * oPalIn.Entries - 1)
        Call CopyMemory(aPalette(0), ByVal oPalIn.lpPalette, 4 * oPalIn.Entries)

        For y = 0 To H
            For x = 0 To W

                '-- Get palette index
                xIn = 4 * x
                Call oPalIn.ClosestIndex(aBitsIn(xIn + 2, y), _
                                         aBitsIn(xIn + 1, y), _
                                         aBitsIn(xIn + 0, y), _
                                         aEntry)
                lEntry = 4 * aEntry

                '-- Set target index/color
                Select Case oDIBOut.BPP
                
                    Case [01_bpp]
                        
                        xIn = x \ 8
                        If (aEntry = 0) Then
                            aBitsOut(xIn, y) = aBitsOut(xIn, y) And Not m_Pow2(7 - (x Mod 8))
                          Else
                            aBitsOut(xIn, y) = aBitsOut(xIn, y) Or m_Pow2(7 - (x Mod 8))
                        End If
                    
                    Case [04_bpp]
                        
                        xIn = x \ 2
                        If (x Mod 2 = 0) Then
                            aBitsOut(xIn, y) = (aBitsOut(xIn, y) And &HF) Or aEntry * &H10
                          Else
                            aBitsOut(xIn, y) = (aBitsOut(xIn, y) And &HF0) Or aEntry
                        End If
                    
                    Case [08_bpp]
                        
                        aBitsOut(x, y) = aEntry
                        
                    Case [24_bpp]
                        
                        xIn = 3 * x
                        aBitsOut(xIn + 0, y) = aPalette(lEntry + 0)
                        aBitsOut(xIn + 1, y) = aPalette(lEntry + 1)
                        aBitsOut(xIn + 2, y) = aPalette(lEntry + 2)

                    Case [32_bpp]
                        
                        xIn = 4 * x
                        aBitsOut(xIn + 0, y) = aPalette(lEntry + 0)
                        aBitsOut(xIn + 1, y) = aPalette(lEntry + 1)
                        aBitsOut(xIn + 2, y) = aPalette(lEntry + 2)
                End Select

                '-- Diffuse error
                If (Diffuse) Then
                    
                    xIn = 4 * x
                    errB = CLng(aBitsIn(xIn + 0, y)) - aPalette(lEntry + 0)
                    errG = CLng(aBitsIn(xIn + 1, y)) - aPalette(lEntry + 1)
                    errR = CLng(aBitsIn(xIn + 2, y)) - aPalette(lEntry + 2)
                    aBitsIn(xIn + 0, y) = aPalette(lEntry + 0)
                    aBitsIn(xIn + 1, y) = aPalette(lEntry + 1)
                    aBitsIn(xIn + 2, y) = aPalette(lEntry + 2)

                    '-- Floyd-Steinberg error diffusion...
                    If (Abs(errB) + Abs(errG) + Abs(errR) > 3) Then
                        If (x < W) Then
                            xIn = 4 * x + 4
                            newB = aBitsIn(xIn + 0, y) + (7 * errB) / 16
                            newG = aBitsIn(xIn + 1, y) + (7 * errG) / 16
                            newR = aBitsIn(xIn + 2, y) + (7 * errR) / 16
                            If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                            If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                            If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                            aBitsIn(xIn + 0, y) = newB
                            aBitsIn(xIn + 1, y) = newG
                            aBitsIn(xIn + 2, y) = newR
                        End If
                        If (y < H) Then
                            xIn = 4 * x
                            yIn = y + 1
                            newB = aBitsIn(xIn + 0, yIn) + (5 * errB) / 16
                            newG = aBitsIn(xIn + 1, yIn) + (5 * errG) / 16
                            newR = aBitsIn(xIn + 2, yIn) + (5 * errR) / 16
                            If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                            If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                            If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                            aBitsIn(xIn + 0, yIn) = newB
                            aBitsIn(xIn + 1, yIn) = newG
                            aBitsIn(xIn + 2, yIn) = newR
                            If (x < W) Then
                                xIn = 4 * x + 4
                                newB = aBitsIn(xIn + 0, yIn) + errB / 16
                                newG = aBitsIn(xIn + 1, yIn) + errG / 16
                                newR = aBitsIn(xIn + 2, yIn) + errR / 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                aBitsIn(xIn + 0, yIn) = newB
                                aBitsIn(xIn + 1, yIn) = newG
                                aBitsIn(xIn + 2, yIn) = newR
                            End If
                            If (x > 0) Then
                                xIn = 4 * x - 4
                                newB = aBitsIn(xIn + 0, yIn) + (3 * errB) / 16
                                newG = aBitsIn(xIn + 1, yIn) + (3 * errG) / 16
                                newR = aBitsIn(xIn + 2, yIn) + (3 * errR) / 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                aBitsIn(xIn + 0, yIn) = newB
                                aBitsIn(xIn + 1, yIn) = newG
                                aBitsIn(xIn + 2, yIn) = newR
                            End If
                        End If
                    End If
                  Else
                    aBitsIn(xIn + 0, y) = aPalette(lEntry + 0)
                    aBitsIn(xIn + 1, y) = aPalette(lEntry + 1)
                    aBitsIn(xIn + 2, y) = aPalette(lEntry + 2)
                End If
            Next x
        Next y
        Call pvUnmapDIB(aBitsIn())
        Call pvUnmapDIB(aBitsOut())
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvMapDIB(oDIB As cDIB, aBits() As Byte, uSA As SAFEARRAY2D)
    
    With uSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB.BytesPerScanline
        .pvData = oDIB.lpBits
    End With
    Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(uSA), 4)
End Sub

Private Sub pvUnmapDIB(aBits() As Byte)

    Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
End Sub

