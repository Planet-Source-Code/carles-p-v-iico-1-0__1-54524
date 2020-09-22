Attribute VB_Name = "mGDIp"
' From great stuff:
'
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
'   by Avery
'
'   Platform SDK Redistributable: GDI+ RTM
'   http://www.microsoft.com/downloads/release.asp?releaseid=32738

Option Explicit

Public Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Public Enum GpStatus
    [OK] = 0
    [GenericError] = 1
    [InvalidParameter] = 2
    [OutOfMemory] = 3
    [ObjectBusy] = 4
    [InsufficientBuffer] = 5
    [NotImplemented] = 6
    [Win32Error] = 7
    [WrongState] = 8
    [Aborted] = 9
    [FileNotFound] = 10
    [ValueOverflow ] = 11
    [AccessDenied] = 12
    [UnknownImageFormat] = 13
    [FontFamilyNotFound] = 14
    [FontStyleNotFound] = 15
    [NotTrueTypeFont] = 16
    [UnsupportedGdiplusVersion] = 17
    [GdiplusNotInitialized ] = 18
    [PropertyNotFound] = 19
    [PropertyNotSupported] = 20
End Enum

'//

Public Enum GpUnit
    [UnitWorld]
    [UnitDisplay]
    [UnitPixel]
    [UnitPoint]
    [UnitInch]
    [UnitDocument]
    [UnitMillimeter]
End Enum

Public Enum ImageLockMode
    [ImageLockModeRead] = &H1
    [ImageLockModeWrite] = &H2
    [ImageLockModeUserInputBuf] = &H4
End Enum

Public Enum ColorMatrixFlags
    [ColorMatrixFlagsDefault] = 0
    [ColorMatrixFlagsSkipGrays]
    [ColorMatrixFlagsAltGray]
End Enum

Public Enum ColorAdjustType
    [ColorAdjustTypeDefault] = 0
    [ColorAdjustTypeBitmap]
    [ColorAdjustTypeBrush]
    [ColorAdjustTypePen]
    [ColorAdjustTypeText]
    [ColorAdjustTypeCount]
    [ColorAdjustTypeAny]
End Enum

Public Enum InterpolationMode
    [InterpolationModeInvalid] = -1
    [InterpolationModeDefault]
    [InterpolationModeLowQuality]
    [InterpolationModeHighQuality]
    [InterpolationModeBilinear]
    [InterpolationModeBicubic]
    [InterpolationModeNearestNeighbor]
    [InterpolationModeHighQualityBilinear]
    [InterpolationModeHighQualityBicubic]
End Enum

Public Enum PixelOffsetMode
    [PixelOffsetModeInvalid] = -1
    [PixelOffsetModeDefault]
    [PixelOffsetModeHighSpeed]
    [PixelOffsetModeHighQuality]
    [PixelOffsetModeNone]
    [PixelOffsetModeHalf]
End Enum

Public Enum QualityMode
    [QualityModeInvalid] = -1
    [QualityModeDefault]
    [QualityModeLow]
    [QualityModeHigh]
End Enum

Public Enum SmoothingMode
    [SmoothingModeInvalid] = [QualityModeInvalid]
    [SmoothingModeDefault] = [QualityModeDefault]
    [SmoothingModeHighSpeed] = [QualityModeLow]
    [SmoothingModeHighQuality] = [QualityModeHigh]
    [SmoothingModeNone]
    [SmoothingModeAntiAlias]
End Enum

Public Enum CompositingMode
    [CompositingModeSourceOver] = 0
    [CompositingModeSourceCopy]
End Enum

Public Enum CompositingQuality
    [CompositingQualityInvalid] = [QualityModeInvalid]
    [CompositingQualityDefault] = [QualityModeDefault]
    [CompositingQualityHighSpeed] = [QualityModeLow]
    [CompositingQualityHighQuality] = [QualityModeHigh]
    [CompositingQualityGammaCorrected]
    [CompositingQualityAssumeLinear]
End Enum

Public Enum PenAlignment
    [PenAlignmentCenter] = 0
    [PenAlignmentInset]
End Enum

Public Enum FillMode
    [FillModeAlternate] = 0
    [FillModeWinding]
End Enum

Public Enum HatchStyle
    [HatchStyleHorizontal] = 0
    [HatchStyleVertical] = 1
    [HatchStyleForwardDiagonal] = 2
    [HatchStyleBackwardDiagonal] = 3
    [HatchStyleCross] = 4
    [HatchStyleDiagonalCross] = 5
    [HatchStyle05Percent] = 6
    [HatchStyle10Percent] = 7
    [HatchStyle20Percent] = 8
    [HatchStyle25Percent] = 9
    [HatchStyle30Percent] = 10
    [HatchStyle40Percent] = 11
    [HatchStyle50Percent] = 12
    [HatchStyle60Percent] = 13
    [HatchStyle70Percent] = 14
    [HatchStyle75Percent] = 15
    [HatchStyle80Percent] = 16
    [HatchStyle90Percent] = 17
    [HatchStyleLightDownwardDiagonal] = 18
    [HatchStyleLightUpwardDiagonal] = 19
    [HatchStyleDarkDownwardDiagonal] = 20
    [HatchStyleDarkUpwardDiagonal] = 21
    [HatchStyleWideDownwardDiagonal] = 22
    [HatchStyleWideUpwardDiagonal] = 23
    [HatchStyleLightVertical] = 24
    [HatchStyleLightHorizontal] = 24
    [HatchStyleNarrowVertical] = 26
    [HatchStyleNarrowHorizontal] = 27
    [HatchStyleDarkVertical] = 28
    [HatchStyleDarkHorizontal ] = 29
    [HatchStyleDashedDownwardDiagonal] = 30
    [HatchStyleDashedUpwardDiagonal] = 31
    [HatchStyleDashedHorizontal] = 32
    [HatchStyleDashedVertical] = 33
    [HatchStyleSmallConfetti] = 34
    [HatchStyleLargeConfetti] = 35
    [HatchStyleZigZag] = 36
    [HatchStyleWave] = 37
    [HatchStyleDiagonalBrick] = 38
    [HatchStyleHorizontalBrick] = 39
    [HatchStyleWeave] = 40
    [HatchStylePlaid] = 41
    [HatchStyleDivot] = 42
    [HatchStyleDottedGrid] = 43
    [HatchStyleDottedDiamond] = 44
    [HatchStyleShingle] = 45
    [HatchStyleTrellis] = 46
    [HatchStyleSphere] = 47
    [HatchStyleSmallGrid] = 48
    [HatchStyleSmallCheckerBoard] = 49
    [HatchStyleLargeCheckerBoard] = 50
    [HatchStyleOutlinedDiamond] = 51
    [HatchStyleSolidDiamond] = 52
    [HatchStyleTotal] = 53
    [HatchStyleLargeGrid] = [HatchStyleCross]
    [HatchStyleMin] = [HatchStyleHorizontal]
    [HatchStyleMax] = [HatchStyleTotal] - 1
End Enum

Public Enum LineCap
    [LineCapFlat] = 0
    [LineCapSquare] = 1
    [LineCapRound] = 2
    [LineCapTriangle] = 3
    [LineCapNoAnchor] = &H10      ' corresponds to flat cap
    [LineCapSquareAnchor] = &H11  ' corresponds to square cap
    [LineCapRoundAnchor] = &H12   ' corresponds to round cap
    [LineCapDiamondAnchor] = &H13 ' corresponds to triangle cap
    [LineCapArrowAnchor] = &H14   ' no correspondence
    [LineCapCustom] = &HFF        ' custom cap
    [LineCapAnchorMask] = &HF0    ' mask to check for anchor or not.
End Enum

Public Enum LineJoin
    [LineJoinMiter] = 0
    [LineJoinBevel]
    [LineJoinRound]
    [LineJoinMiterClipped]
End Enum

Public Enum FontStyle
    [FontStyleRegular] = 0
    [FontStyleBold] = 1
    [FontStyleItalic] = 2
    [FontStyleBoldItalic] = 3
    [FontStyleUnderline] = 4
    [FontStyleStrikeout] = 8
End Enum

Public Enum TextRenderingHint
    [TextRenderingHintSystemDefault] = 0        ' Glyph with system default rendering hint
    [TextRenderingHintSingleBitPerPixelGridFit] ' Glyph bitmap with hinting
    [TextRenderingHintSingleBitPerPixel]        ' Glyph bitmap without hinting
    [TextRenderingHintAntiAliasGridFit]         ' Glyph anti-alias bitmap with hinting
    [TextRenderingHintAntiAlias]                ' Glyph anti-alias bitmap without hinting
    [TextRenderingHintClearTypeGridFit]         ' Glyph CT bitmap with hinting
End Enum

Public Type BITMAPDATA
    Width       As Long
    Height      As Long
    Stride      As Long
    PixelFormat As Long
    Scan0       As Long
    Reserved    As Long
End Type

Public Type COLORMATRIX
    m(0 To 4, 0 To 4) As Single
End Type

Public Type POINTL
    x As Long
    y As Long
End Type

Public Type POINTF
    x As Single
    y As Single
End Type

Public Type RECTL
    x As Long
    y As Long
    W As Long
    H As Long
End Type

Public Type RECTF
    x As Single
    y As Single
    W As Single
    H As Single
End Type

Public Const PixelFormat1bppIndexed As Long = &H30101
Public Const PixelFormat4bppIndexed As Long = &H30402
Public Const PixelFormat8bppIndexed As Long = &H30803
Public Const PixelFormat24bppRGB    As Long = &H21808
Public Const PixelFormat32bppRGB    As Long = &H22009
Public Const PixelFormat32bppARGB   As Long = &H26200A

Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, InputBuf As GdiplusStartupInput, Optional ByVal OutputBuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus

Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal Filename As String, hImage As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal hImage As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal hImage As Long, Height As Long) As GpStatus

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As GpStatus
Public Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal hImage As Long, hGraphics As Long) As GpStatus
Public Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal hGraphics As Long, ByVal lColor As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GpStatus

Public Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Any, hBitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As Any, gdiBitmapData As Any, hBitmap As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GpStatus

Public Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As InterpolationMode) As GpStatus
Public Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As PixelOffsetMode) As GpStatus
Public Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As SmoothingMode) As GpStatus
Public Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As CompositingMode) As GpStatus
Public Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal hGraphics As Long, ByVal Quality As CompositingQuality) As GpStatus
Public Declare Function GdipCreateImageAttributes Lib "gdiplus" (hAttributes As Long) As GpStatus
Public Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal hAttributes As Long, ByVal ColorAdjust As ColorAdjustType, ByVal EnableFlag As Long, Matrix As COLORMATRIX, GrayMatrix As Any, ByVal Flags As ColorMatrixFlags) As GpStatus
Public Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal hAttributes As Long) As GpStatus

Public Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, rect As RECTL, ByVal Flags As Long, ByVal PixelFormat As Long, LockedBitmapData As BITMAPDATA) As GpStatus
Public Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, LockedBitmapData As BITMAPDATA) As GpStatus
Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus

Public Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal Unit As GpUnit, hPen As Long) As GpStatus
Public Declare Function GdipCreatePen2 Lib "gdiplus" (ByVal hBrush As Long, ByVal Width As Single, ByVal Unit As GpUnit, hPen As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal hPen As Long) As GpStatus

Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal Color As Long, hBrush As Long) As GpStatus
Public Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal Style As HatchStyle, ByVal ForeColor As Long, ByVal BackColor As Long, hBrush As Long) As GpStatus
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As GpStatus

Public Declare Function GdipSetPenMode Lib "gdiplus" (ByVal hPen As Long, ByVal Mode As PenAlignment) As GpStatus
Public Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal hPen As Long, ByVal StartCap As LineCap) As GpStatus
Public Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal hPen As Long, ByVal EndCap As LineCap) As GpStatus
Public Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal hPen As Long, ByVal Join As LineJoin) As GpStatus

Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal hBitmap As Long, ByVal x As Long, ByVal y As Long, lColor As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal hBitmap As Long, ByVal x As Long, ByVal y As Long, ByVal lColor As Long) As GpStatus

Public Declare Function GdipDrawLineI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As GpStatus
Public Declare Function GdipDrawLine Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As GpStatus
Public Declare Function GdipDrawLinesI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, Points As POINTL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawLines Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, Points As POINTF, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipFillRectangle Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipFillEllipse Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawArc Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal StartAngle As Single, ByVal SweepAngle As Single) As GpStatus
Public Declare Function GdipDrawArcI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hPen As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal StartAngle As Single, ByVal SweepAngle As Single) As GpStatus

Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As String, ByVal FontCollection As Long, FontFamily As Long) As GpStatus
Public Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal hFontFamily As Long) As GpStatus
Public Declare Function GdipCreateFont Lib "gdiplus" (ByVal hFontFamily As Long, ByVal emSize As Single, ByVal Style As FontStyle, ByVal Unit As GpUnit, hCreatedFont As Long) As GpStatus
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal hCurFont As Long) As GpStatus
Public Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As TextRenderingHint) As GpStatus
Public Declare Function GdipMeasureString Lib "gdiplus" (ByVal hGraphics As Long, ByVal Str As String, ByVal Length As Long, ByVal TheFont As Long, LayoutRect As RECTF, ByVal StringFormat As Long, BoundingBox As RECTF, CodePointsFitted As Long, LinesFilled As Long) As GpStatus
Public Declare Function GdipDrawString Lib "gdiplus" (ByVal hGraphics As Long, ByVal Str As String, ByVal Length As Long, ByVal TheFont As Long, LayoutRect As RECTF, ByVal StringFormat As Long, ByVal hBrush As Long) As GpStatus

'//

Public Sub ARGBBitmapFromGDIDIB32(oDIB As cDIB, hBitmap As Long)

  Dim bmpRect As RECTL
  Dim bmpData As BITMAPDATA
        
    '-- Prepare image info
    With bmpRect
        .W = oDIB.Width
        .H = oDIB.Height
    End With
    With bmpData
        .Width = oDIB.Width
        .Height = oDIB.Height
        .Stride = -oDIB.BytesPerScanline
        .PixelFormat = [PixelFormat32bppARGB]
        .Scan0 = oDIB.lpBits - .Stride * (oDIB.Height - 1) ' Vertical flip
    End With
    
    '-- Create a blank ARGB GDI+ bitmap
    Call GdipCreateBitmapFromScan0(oDIB.Width, oDIB.Height, 0, [PixelFormat32bppARGB], ByVal 0, hBitmap)
    '-- Lock bits and assign color data
    Call GdipBitmapLockBits(hBitmap, bmpRect, [ImageLockModeWrite] Or [ImageLockModeUserInputBuf], [PixelFormat32bppARGB], bmpData)
    Call GdipBitmapUnlockBits(hBitmap, bmpData)
End Sub

Public Sub ARGBBitmapToGDIDIB32(oDIB As cDIB, hBitmap As Long)
  
  Dim bmpRect As RECTL
  Dim bmpData As BITMAPDATA
        
    '-- Prepare image info
    With bmpRect
        .W = oDIB.Width
        .H = oDIB.Height
    End With
    With bmpData
        .Width = oDIB.Width
        .Height = oDIB.Height
        .Stride = -oDIB.BytesPerScanline
        .PixelFormat = [PixelFormat32bppARGB]
        .Scan0 = oDIB.lpBits - .Stride * (oDIB.Height - 1) ' Vertical flip
    End With
    
    '-- Lock bits and assign color data
    Call GdipBitmapLockBits(hBitmap, bmpRect, [ImageLockModeRead] Or [ImageLockModeUserInputBuf], [PixelFormat32bppARGB], bmpData)
    Call GdipBitmapUnlockBits(hBitmap, bmpData)
End Sub

Public Sub BitmapToGDIDIB32(oDIB As cDIB, hBitmap As Long)
  
  Dim bmpRect   As RECTL
  Dim bmpData   As BITMAPDATA
  Dim hBitmap32 As Long, hGraphics As Long
    
    '-- Create a blank ARGB GDI+ bitmap
    Call GdipCreateBitmapFromScan0(oDIB.Width, oDIB.Height, 0, [PixelFormat32bppARGB], ByVal 0, hBitmap32)
    Call GdipGetImageGraphicsContext(hBitmap32, hGraphics)
    
    '-- Draw (translate) source image onto ARGB bitmap
    Call GdipDrawImageRectI(hGraphics, hBitmap, 0, 0, oDIB.Width, oDIB.Height)
    
    '-- Prepare image info
    With bmpRect
        .W = oDIB.Width
        .H = oDIB.Height
    End With
    With bmpData
        .Width = oDIB.Width
        .Height = oDIB.Height
        .Stride = -oDIB.BytesPerScanline
        .PixelFormat = [PixelFormat32bppARGB]
        .Scan0 = oDIB.lpBits - .Stride * (oDIB.Height - 1) ' Vertical flip
    End With
    
    '-- Lock bits and assign color data
    Call GdipBitmapLockBits(hBitmap32, bmpRect, [ImageLockModeRead] Or [ImageLockModeUserInputBuf], [PixelFormat32bppARGB], bmpData)
    Call GdipBitmapUnlockBits(hBitmap32, bmpData)
    
    '-- Clean up
    Call GdipDisposeImage(hBitmap32)
    Call GdipDeleteGraphics(hGraphics)
End Sub

Public Function StretchDIB32(oDIB As cDIB, _
                ByVal hDC As Long, _
                ByVal x As Long, ByVal y As Long, _
                ByVal nWidth As Long, ByVal nHeight As Long, _
                Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, _
                Optional ByVal nSrcWidth As Long, Optional ByVal nSrcHeight As Long, _
                Optional ByVal Alpha As Byte = 255, _
                Optional ByVal Blend As Boolean = True, _
                Optional ByVal Interpolate As Boolean = False) As Long

  Dim gplRet As Long
  
  Dim hGraphics   As Long
  Dim hAttributes As Long
  Dim uMatrix     As COLORMATRIX
  Dim hBitmap     As Long
  Dim bmpRect     As RECTL
  Dim bmpData     As BITMAPDATA
  
    If (oDIB.BPP = 32) Then
        
        If (nSrcWidth = 0) Then nSrcWidth = oDIB.Width
        If (nSrcHeight = 0) Then nSrcHeight = oDIB.Height
      
        With bmpRect
            .W = oDIB.Width
            .H = oDIB.Height
        End With
        
        With bmpData
            .Width = oDIB.Width
            .Height = oDIB.Height
            .Stride = -oDIB.BytesPerScanline
            .PixelFormat = [PixelFormat32bppARGB]
            .Scan0 = oDIB.lpBits - .Stride * (oDIB.Height - 1) ' Vertical flip
        End With
        
        '-- Initialize Graphics object
        gplRet = GdipCreateFromHDC(hDC, hGraphics)
        
        '-- Initialize blank Bitmap and assign GDI DIB data
        gplRet = GdipCreateBitmapFromScan0(oDIB.Width, oDIB.Height, 0, [PixelFormat32bppARGB], ByVal 0, hBitmap)
        gplRet = GdipBitmapLockBits(hBitmap, bmpRect, [ImageLockModeWrite] Or [ImageLockModeUserInputBuf], [PixelFormat32bppARGB], bmpData)
        gplRet = GdipBitmapUnlockBits(hBitmap, bmpData)

        '-- Prepare/Set image attributes (global alpha)
        With uMatrix
            .m(0, 0) = 1
            .m(1, 1) = 1
            .m(2, 2) = 1
            .m(3, 3) = Alpha / 255
            .m(4, 4) = 1
        End With
        gplRet = GdipCreateImageAttributes(hAttributes)
        gplRet = GdipSetImageAttributesColorMatrix(hAttributes, [ColorAdjustTypeDefault], True, uMatrix, ByVal 0, [ColorMatrixFlagsDefault])
        
        '-- Draw ARGB
        gplRet = GdipSetCompositingMode(hGraphics, [CompositingModeSourceOver] * -(Not Blend))
        gplRet = GdipSetInterpolationMode(hGraphics, [InterpolationModeNearestNeighbor] + -(2 * Interpolate))
        gplRet = GdipSetPixelOffsetMode(hGraphics, [PixelOffsetModeHighQuality])
        gplRet = GdipDrawImageRectRectI(hGraphics, hBitmap, x, y, nWidth, nHeight, xSrc, ySrc, nSrcWidth, nSrcHeight, [UnitPixel], hAttributes)
        
        '-- Clean up
        gplRet = GdipDeleteGraphics(hGraphics)
        gplRet = GdipDisposeImage(hBitmap)
        gplRet = GdipDisposeImageAttributes(hAttributes)
        
        '-- Success
        StretchDIB32 = (gplRet = [OK])
    End If
End Function


