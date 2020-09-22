VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIconCanvasEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cIconCanvasEx.cls
' Author:        Carles P.V.
' Dependencies:  from ucIconCanvas.ctl
' Last revision: 2004.06.14
'================================================

Option Explicit

'-- API:

Private Type ARGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type HSL
    H As Long
    S As Long
    L As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Type LOGFONT
   lfHeight         As Long
   lfWidth          As Long
   lfEscapement     As Long
   lfOrientation    As Long
   lfWeight         As Long
   lfItalic         As Byte
   lfUnderline      As Byte
   lfStrikeOut      As Byte
   lfCharSet        As Byte
   lfOutPrecision   As Byte
   lfClipPrecision  As Byte
   lfQuality        As Byte
   lfPitchAndFamily As Byte
   lfFaceName(32)   As Byte
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims          As Integer
    fFeatures      As Integer
    cbElements     As Long
    cLocks         As Long
    pvData         As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Const PS_GEOMETRIC           As Long = &H10000
Private Const PS_JOIN_MITER          As Long = &H2000
Private Const PS_ENDCAP_SQUARE       As Long = &H100
Private Const PS_ENDCAP_ROUND        As Long = &H0
Private Const PS_JOIN_ROUND          As Long = &H0

Private Const BS_SOLID               As Long = 0
Private Const BS_NULL                As Long = 1

Private Const LOGPIXELSY             As Long = 90
Private Const FW_NORMAL              As Long = 400
Private Const FW_BOLD                As Long = 700
Private Const FF_DONTCARE            As Long = 0
Private Const DEFAULT_QUALITY        As Long = 0
Private Const DEFAULT_PITCH          As Long = 0
Private Const DEFAULT_CHARSET        As Long = 1
Private Const NONANTIALIASED_QUALITY As Long = 3
Private Const DT_CALCRECT            As Long = &H400
Private Const DT_NOCLIP              As Long = &H100

Private Const TRANSPARENT            As Long = 1

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long

Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDst As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long

'//

'-- Public Enums.:
Private Enum icoColorSample
    [icoSampleA] = 0
    [icoSampleB]
End Enum

Public Enum icoLayerCts
    [icoXORSolid] = 0
    [icoXORAlpha] = 1
    [icoAND] = 2
End Enum

Public Enum icoPixelMaskCts
    [icoAll]
    [icoEven]
    [icoOdd]
End Enum

Public Enum icoShapeCts
    [icoRectangle] = 0
    [icoRectangleSolid]
    [icoEllipse]
    [icoEllipseSolid]
End Enum

'-- Property Variables:
Private m_ColorScreen       As Long
Private m_IsAScreen         As Boolean
Private m_IsBScreen         As Boolean
Private m_ColorLngA         As Long
Private m_ColorLngB         As Long
Private m_ColorIdxA         As Byte
Private m_ColorIdxB         As Byte
Private m_SwapColors        As Boolean
Private m_PixelMask         As icoPixelMaskCts
Private m_StraightLineWidth As Long
Private m_BrushLineWidth    As Long
Private m_ShapeLineWidth    As Long
Private m_Shape             As icoShapeCts
Private m_Alpha             As Byte
Private m_FillToleranceA    As Byte
Private m_FillToleranceH    As Long
Private m_FillToleranceS    As Long
Private m_FillToleranceL    As Long
Private m_Antialias         As Boolean
Private m_oFont             As StdFont
Private m_Text              As String
Private m_MaskImport        As Boolean
Private m_MaskExport        As Boolean
Private m_Import32bppOpaque As Boolean

'-- Private Variables:
Private m_Pow2(31) As Long
Private m_PXOFF1   As Single
Private m_PXOFF2   As Single



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Initialize()

  Dim lIdx As Long
    
    '-- Set default values
    m_IsAScreen = False
    m_IsBScreen = False
    m_SwapColors = False
    m_ColorLngA = &H0
    m_ColorLngB = &HFFFFFF
    m_ColorIdxA = 0
    m_ColorIdxB = 1
    m_PixelMask = [icoAll]
    m_StraightLineWidth = 1
    m_BrushLineWidth = 1
    m_Shape = [icoRectangle]
    m_Alpha = 255
    m_FillToleranceA = 0
    m_FillToleranceH = 0
    m_FillToleranceS = 0
    m_FillToleranceL = 0
    m_Antialias = False
    Set m_oFont = New StdFont
    m_Text = vbNullString
    m_MaskImport = False
    m_MaskExport = False
    m_Import32bppOpaque = False
    
    '-- Quick 2^n
    For lIdx = 0 To 30
        m_Pow2(lIdx) = 2 ^ lIdx
    Next lIdx
    m_Pow2(31) = &H80000000
    m_PXOFF1 = 0.448
    m_PXOFF2 = 0.449
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get ColorScreen() As Long
    ColorScreen = m_ColorScreen
End Property
Public Property Let ColorScreen(ByVal New_ColorScreen As Long)
    m_ColorScreen = New_ColorScreen
End Property

Public Property Let IsAScreen(ByVal New_IsAScreen As Boolean)
    m_IsAScreen = New_IsAScreen
End Property
Public Property Get IsAScreen() As Boolean
    IsAScreen = m_IsAScreen
End Property

Public Property Let IsBScreen(ByVal New_IsBScreen As Boolean)
    m_IsBScreen = New_IsBScreen
End Property
Public Property Get IsBScreen() As Boolean
    IsBScreen = m_IsBScreen
End Property

Public Property Let ColorLngA(ByVal New_ColorLngA As Long)
    m_ColorLngA = New_ColorLngA
End Property
Public Property Get ColorLngA() As Long
    ColorLngA = m_ColorLngA
End Property

Public Property Let ColorLngB(ByVal New_ColorLngB As Long)
    m_ColorLngB = New_ColorLngB
End Property
Public Property Get ColorLngB() As Long
    ColorLngB = m_ColorLngB
End Property

Public Property Let ColorIdxA(ByVal New_ColorIdxA As Byte)
    m_ColorIdxA = New_ColorIdxA
End Property
Public Property Get ColorIdxA() As Byte
    ColorIdxA = m_ColorIdxA
End Property

Public Property Let ColorIdxB(ByVal New_ColorIdxB As Byte)
    m_ColorIdxB = New_ColorIdxB
End Property
Public Property Get ColorIdxB() As Byte
    ColorIdxB = m_ColorIdxB
End Property

Public Property Let SwapColors(ByVal New_SwapColors As Boolean)
    m_SwapColors = New_SwapColors
End Property
Public Property Get SwapColors() As Boolean
    SwapColors = m_SwapColors
End Property

Public Property Let StraightLineWidth(ByVal New_StraightLineWidth As Long)
    m_StraightLineWidth = New_StraightLineWidth
End Property
Public Property Get StraightLineWidth() As Long
    StraightLineWidth = m_StraightLineWidth
End Property

Public Property Let PixelMask(ByVal New_PixelMask As icoPixelMaskCts)
    m_PixelMask = New_PixelMask
End Property
Public Property Get PixelMask() As icoPixelMaskCts
    PixelMask = m_PixelMask
End Property

Public Property Let BrushLineWidth(ByVal New_BrushLineWidth As Long)
    m_BrushLineWidth = New_BrushLineWidth
End Property
Public Property Get BrushLineWidth() As Long
    BrushLineWidth = m_BrushLineWidth
End Property

Public Property Let ShapeLineWidth(ByVal New_ShapeLineWidth As Long)
    m_ShapeLineWidth = New_ShapeLineWidth
End Property
Public Property Get ShapeLineWidth() As Long
    ShapeLineWidth = m_ShapeLineWidth
End Property

Public Property Let Shape(ByVal New_Shape As icoShapeCts)
    m_Shape = New_Shape
End Property
Public Property Get Shape() As icoShapeCts
    Shape = m_Shape
End Property

Public Property Let Alpha(ByVal New_Alpha As Byte)
    m_Alpha = New_Alpha
End Property
Public Property Get Alpha() As Byte
    Alpha = m_Alpha
End Property

Public Property Let FillToleranceA(ByVal New_FillToleranceA As Byte)
    m_FillToleranceA = New_FillToleranceA
End Property
Public Property Get FillToleranceA() As Byte
    FillToleranceA = m_FillToleranceA
End Property

Public Property Let FillToleranceH(ByVal New_FillToleranceH As Long)
    m_FillToleranceH = New_FillToleranceH
End Property
Public Property Get FillToleranceH() As Long
    FillToleranceH = m_FillToleranceH
End Property

Public Property Let FillToleranceS(ByVal New_FillToleranceS As Long)
    m_FillToleranceS = New_FillToleranceS
End Property
Public Property Get FillToleranceS() As Long
    FillToleranceS = m_FillToleranceS
End Property

Public Property Let FillToleranceL(ByVal New_FillToleranceL As Long)
    m_FillToleranceL = New_FillToleranceL
End Property
Public Property Get FillToleranceL() As Long
    FillToleranceL = m_FillToleranceL
End Property

Public Property Let Antialias(ByVal New_Antialias As Boolean)
    m_Antialias = New_Antialias
End Property
Public Property Get Antialias() As Boolean
    Antialias = m_Antialias
End Property

Public Property Let Font(ByVal New_Font As StdFont)
    Set m_oFont = New_Font
End Property
Public Property Get Font() As StdFont
    Set Font = m_oFont
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
End Property
Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let MaskImport(ByVal New_MaskImport As Boolean)
    m_MaskImport = New_MaskImport
End Property
Public Property Get MaskImport() As Boolean
    MaskImport = m_MaskImport
End Property

Public Property Let MaskExport(ByVal New_MaskExport As Boolean)
    m_MaskExport = New_MaskExport
End Property
Public Property Get MaskExport() As Boolean
    MaskExport = m_MaskExport
End Property

Public Property Let Import32bppOpaque(ByVal New_Import32bppOpaque As Boolean)
    m_Import32bppOpaque = New_Import32bppOpaque
End Property
Public Property Get Import32bppOpaque() As Boolean
    Import32bppOpaque = Import32bppOpaque
End Property

'========================================================================================
' Friend methods
'========================================================================================

Friend Sub DrawPixels(oDIB As cDIB, _
           Points() As POINTL, _
           ByVal Layer As icoLayerCts)
            
  Dim hBitmap   As Long, hTmpBitmap As Long
  Dim hGraphics As Long
  Dim lColorPen As Long
  Dim lIdx      As Long
  Dim bPaint    As Boolean
  Dim rctBounds As RECT2
  
    If (Layer = [icoXORAlpha]) Then
    
        Call SetRect(rctBounds, 0, 0, oDIB.Width, oDIB.Height)
        
        Call GdipCreateBitmapFromScan0(oDIB.Width, oDIB.Height, 0, [PixelFormat32bppARGB], ByVal 0, hTmpBitmap)
        Call ARGBBitmapFromGDIDIB32(oDIB, hBitmap)
        Call GdipGetImageGraphicsContext(hBitmap, hGraphics)
        
        lColorPen = pvColorRef(oDIB, Layer, [icoSampleA], m_Alpha)
        For lIdx = 0 To UBound(Points())
            Select Case m_PixelMask
                Case [icoAll]:  bPaint = True
                Case [icoEven]: bPaint = (Points(lIdx).x + Points(lIdx).y) Mod 2 = 1
                Case [icoOdd]:  bPaint = (Points(lIdx).x + Points(lIdx).y) Mod 2 = 0
            End Select
            If (bPaint And PtInRect(rctBounds, Int(Points(lIdx).x), Int(Points(lIdx).y))) Then
                Call GdipBitmapSetPixel(hTmpBitmap, Points(lIdx).x, Points(lIdx).y, lColorPen)
            End If
        Next lIdx
        Call GdipDrawImageRectI(hGraphics, hTmpBitmap, 0, 0, oDIB.Width, oDIB.Height)
        
        Call ARGBBitmapToGDIDIB32(oDIB, hBitmap)
        Call GdipDeleteGraphics(hGraphics)
        Call GdipDisposeImage(hBitmap)
        Call GdipDisposeImage(hTmpBitmap)
        
      Else
      
        lColorPen = pvColorRef(oDIB, Layer, [icoSampleA], 0)
        For lIdx = 0 To UBound(Points())
            Select Case m_PixelMask
                Case [icoAll]:  bPaint = True
                Case [icoEven]: bPaint = (Points(lIdx).x + Points(lIdx).y) Mod 2 = 1
                Case [icoOdd]:  bPaint = (Points(lIdx).x + Points(lIdx).y) Mod 2 = 0
            End Select
            If (bPaint) Then
                Call SetPixelV(oDIB.hDC, Points(lIdx).x, Points(lIdx).y, lColorPen)
            End If
        Next lIdx
    End If
End Sub

Friend Sub DrawStraightLine(oDIB As cDIB, _
           ByVal x1 As Single, ByVal y1 As Single, _
           ByVal x2 As Single, ByVal y2 As Single, _
           ByVal Layer As icoLayerCts)
           
  Dim hBitmap   As Long
  Dim hGraphics As Long
  Dim lColorPen As Long
  Dim uLogBrush As LOGBRUSH
  Dim hPen      As Long
  Dim hOldPen   As Long
  Dim ptFrom    As POINTAPI

    If (Layer = [icoXORAlpha]) Then
    
        Call ARGBBitmapFromGDIDIB32(oDIB, hBitmap)
        Call GdipGetImageGraphicsContext(hBitmap, hGraphics)
        
        lColorPen = pvColorRef(oDIB, Layer, [icoSampleA], m_Alpha)
        Call GdipCreatePen1(lColorPen, m_StraightLineWidth, [UnitPixel], hPen)
        Call GdipSetPenStartCap(hPen, [LineCapRound])
        Call GdipSetPenEndCap(hPen, [LineCapRound])
                
        If (m_Antialias) Then
            m_PXOFF1 = 0.001: m_PXOFF2 = 0.002
          Else
            m_PXOFF1 = 0.448: m_PXOFF2 = 0.449
            x1 = Int(x1): y1 = Int(y1)
            x2 = Int(x2): y2 = Int(y2)
        End If

        If (x1 = x2 And y1 = y2) Then
            x1 = x1 + m_PXOFF1
            y1 = y1 + m_PXOFF1
            x2 = x2 + m_PXOFF2 * -(m_StraightLineWidth > 1)
            y2 = y2 + m_PXOFF2 * -(m_StraightLineWidth > 1)
            Call GdipSetSmoothingMode(hGraphics, [SmoothingModeAntiAlias] * -m_Antialias)
            Call GdipSetPixelOffsetMode(hGraphics, [PixelOffsetModeHalf])
          Else
            Call GdipSetSmoothingMode(hGraphics, [SmoothingModeAntiAlias] * -m_Antialias)
            Call GdipSetPixelOffsetMode(hGraphics, [PixelOffsetModeHalf] - Not m_Antialias)
        End If

        Call GdipDrawLine(hGraphics, hPen, x1, y1, x2, y2)
        
        Call ARGBBitmapToGDIDIB32(oDIB, hBitmap)
        Call GdipDeletePen(hPen)
        Call GdipDeleteGraphics(hGraphics)
        Call GdipDisposeImage(hBitmap)
        
      Else
        
        lColorPen = pvColorRef(oDIB, Layer, [icoSampleA], 0)
        With uLogBrush
            .lbStyle = BS_SOLID
            .lbColor = lColorPen
            .lbHatch = 0
        End With
        hPen = ExtCreatePen(PS_GEOMETRIC + PS_ENDCAP_ROUND + PS_JOIN_ROUND, m_StraightLineWidth, uLogBrush, 0, ByVal 0)
        hOldPen = SelectObject(oDIB.hDC, hPen)
        
        Call MoveToEx(oDIB.hDC, x1, y1, ptFrom)
        Call LineTo(oDIB.hDC, x2, y2)
        Call SetPixelV(oDIB.hDC, x2, y2, lColorPen)
        
        Call SelectObject(oDIB.hDC, hOldPen)
        Call DeleteObject(hPen)
    End If
End Sub

Friend Sub DrawBrushLines(oDIB As cDIB, _
           Points() As POINTF, _
           ByVal Layer As icoLayerCts)
           
  Dim hBitmap   As Long
  Dim hGraphics As Long
  Dim lColorPen As Long
  Dim uLogBrush As LOGBRUSH
  Dim hPen      As Long
  Dim hOldPen   As Long
  Dim ptFrom    As POINTAPI
  Dim uPoints() As POINTF
  Dim lIdx      As Long

    If (Layer = [icoXORAlpha]) Then
    
        Call ARGBBitmapFromGDIDIB32(oDIB, hBitmap)
        Call GdipGetImageGraphicsContext(hBitmap, hGraphics)
        
        lColorPen = pvColorRef(oDIB, Layer, [icoSampleA], m_Alpha)
        Call GdipCreatePen1(lColorPen, m_BrushLineWidth, [UnitPixel], hPen)
        Call GdipSetPenStartCap(hPen, [LineCapRound])
        Call GdipSetPenEndCap(hPen, [LineCapRound])
        Call GdipSetPenLineJoin(hPen, [LineJoinRound])
                
        uPoints() = Points()
        If (m_Antialias) Then
            m_PXOFF1 = 0.001: m_PXOFF2 = 0.002
          Else
            m_PXOFF1 = 0.448: m_PXOFF2 = 0.449
            For lIdx = 0 To UBound(uPoints())
                With uPoints(lIdx)
                    .x = Int(.x)
                    .y = Int(.y)
                End With
            Next lIdx
        End If

        Call GdipSetSmoothingMode(hGraphics, [SmoothingModeAntiAlias] * -m_Antialias)
        Call GdipSetPixelOffsetMode(hGraphics, [PixelOffsetModeHalf])
        
        With uPoints(0)
            .x = .x + m_PXOFF1 * -(m_BrushLineWidth > 1)
            .y = .y + m_PXOFF1 * -(m_BrushLineWidth > 1)
        End With
        For lIdx = 1 To UBound(uPoints())
            With uPoints(lIdx)
                .x = .x + m_PXOFF2
                .y = .y + m_PXOFF2
            End With
        Next lIdx
        
        Call GdipDrawLines(hGraphics, hPen, uPoints(0), UBound(uPoints()) + 1)
        
        Call ARGBBitmapToGDIDIB32(oDIB, hBitmap)
        Call GdipDeletePen(hPen)
        Call GdipDeleteGraphics(hGraphics)
        Call GdipDisposeImage(hBitmap)
        
      Else
        
        lColorPen = pvColorRef(oDIB, Layer, [icoSampleA], 0)
        With uLogBrush
            .lbStyle = BS_SOLID
            .lbColor = lColorPen
            .lbHatch = 0
        End With
        hPen = ExtCreatePen(PS_GEOMETRIC + PS_ENDCAP_ROUND + PS_JOIN_ROUND, m_BrushLineWidth, uLogBrush, 0, ByVal 0)
        hOldPen = SelectObject(oDIB.hDC, hPen)
        
        Call MoveToEx(oDIB.hDC, Points(0).x, Points(0).y, ptFrom)
        For lIdx = 0 To UBound(Points()) - 1
            Call LineTo(oDIB.hDC, Points(lIdx + 1).x, Points(lIdx + 1).y)
        Next lIdx
        Call SetPixelV(oDIB.hDC, Points(lIdx).x, Points(lIdx).y, lColorPen)
        
        Call SelectObject(oDIB.hDC, hOldPen)
        Call DeleteObject(hPen)
    End If
End Sub

Friend Sub DrawShape(oDIB As cDIB, _
           ByVal x1 As Single, ByVal y1 As Single, _
           ByVal x2 As Single, ByVal y2 As Single, _
           ByVal Layer As icoLayerCts)
    
  Dim hBitmap     As Long, hTmpBitmap   As Long
  Dim hGraphics   As Long, hTmpGraphics As Long
  Dim hAttributes As Long
  Dim uMatrix     As COLORMATRIX
  Dim lColorPen   As Long
  Dim lColorBrush As Long
  Dim hPen        As Long
  Dim hOldPen     As Long
  Dim uLogBrush   As LOGBRUSH
  Dim hBrush      As Long
  Dim hOldBrush   As Long
  Dim ptFrom      As POINTAPI
  Dim snW         As Single
  Dim snH         As Single
      
    If (Layer = [icoXORAlpha]) Then
        
        Call ARGBBitmapFromGDIDIB32(oDIB, hBitmap)
        Call GdipGetImageGraphicsContext(hBitmap, hGraphics)
        
        Call GdipCreateBitmapFromScan0(oDIB.Width, oDIB.Height, 0, [PixelFormat32bppARGB], ByVal 0, hTmpBitmap)
        Call GdipGetImageGraphicsContext(hTmpBitmap, hTmpGraphics)
        
        lColorPen = pvColorRef(oDIB, Layer, [icoSampleA], 255)
        Call GdipCreatePen1(lColorPen, m_ShapeLineWidth, [UnitPixel], hPen)
        If (m_Shape > [icoRectangleSolid]) Then
            Call GdipSetPenStartCap(hPen, [LineCapRound])
            Call GdipSetPenEndCap(hPen, [LineCapRound])
            Call GdipSetPenLineJoin(hPen, [LineJoinRound])
          Else
            Call GdipSetPenStartCap(hPen, [LineCapSquare])
            Call GdipSetPenEndCap(hPen, [LineCapSquare])
            Call GdipSetPenLineJoin(hPen, [LineJoinMiter])
        End If
        
        lColorBrush = pvColorRef(oDIB, Layer, [icoSampleA], 255)
        Call GdipCreateSolidFill(lColorBrush, hBrush)
            
        If (m_Antialias And m_Shape > [icoRectangleSolid]) Then
            m_PXOFF1 = 0.001: m_PXOFF2 = 0.002
          Else
            m_PXOFF1 = 0.498: m_PXOFF2 = 0.499
            x1 = Int(x1): y1 = Int(y1)
            x2 = Int(x2): y2 = Int(y2)
        End If
        
        Select Case True
        
            Case x1 = x2 And y1 = y2
                
                Call GdipSetSmoothingMode(hTmpGraphics, [SmoothingModeAntiAlias] * -(m_Antialias And m_Shape > [icoRectangleSolid]))
                Call GdipSetPixelOffsetMode(hTmpGraphics, [PixelOffsetModeHalf] - Not (m_Shape > [icoRectangleSolid]))
                
                x1 = x1 + m_PXOFF1
                y1 = y1 + m_PXOFF1
                x2 = x2 + m_PXOFF2 * -(m_ShapeLineWidth > 1)
                y2 = y2 + m_PXOFF1 * -(m_ShapeLineWidth > 1)
                
                Call GdipDrawLine(hTmpGraphics, hPen, x1, y1, x2, y2)
       
            Case x1 = x2 Xor y1 = y2
                
                Call GdipSetSmoothingMode(hTmpGraphics, [SmoothingModeAntiAlias] * -(m_Antialias And m_Shape > [icoRectangleSolid]))
                Call GdipSetPixelOffsetMode(hTmpGraphics, [PixelOffsetModeHalf] - Not (m_Antialias And m_Shape > [icoRectangleSolid]))
                
                Call GdipDrawLine(hTmpGraphics, hPen, x1, y1, x2, y2)
                
            Case Else
                
                Call GdipSetSmoothingMode(hTmpGraphics, [SmoothingModeAntiAlias] * -(m_Antialias And m_Shape > [icoRectangleSolid]))
                Call GdipSetPixelOffsetMode(hTmpGraphics, [PixelOffsetModeHalf] - Not (m_Antialias And m_Shape > [icoRectangleSolid]))
        
                Call pvCorrectShapeCoordsGDIp(x1, y1, x2, y2)
                snW = x2 - x1
                snH = y2 - y1
                
                Select Case m_Shape
                
                    Case [icoRectangle]
                        Call GdipDrawLine(hTmpGraphics, hPen, x1, y1, x2, y1)
                        Call GdipDrawLine(hTmpGraphics, hPen, x1, y1, x1, y2)
                        Call GdipDrawLine(hTmpGraphics, hPen, x2, y1, x2, y2)
                        Call GdipDrawLine(hTmpGraphics, hPen, x1, y2, x2, y2)
                        
                    Case [icoRectangleSolid]
                        Call GdipFillRectangle(hTmpGraphics, hBrush, x1, y1, snW, snH)
                        Call GdipDrawLine(hTmpGraphics, hPen, x1, y1, x2, y1)
                        Call GdipDrawLine(hTmpGraphics, hPen, x1, y1, x1, y2)
                        Call GdipDrawLine(hTmpGraphics, hPen, x2, y1, x2, y2)
                        Call GdipDrawLine(hTmpGraphics, hPen, x1, y2, x2, y2)
                    
                    Case [icoEllipse]
                        Call GdipDrawArc(hTmpGraphics, hPen, x1, y1, snW, snH, 0, 90)
                        Call GdipDrawArc(hTmpGraphics, hPen, x1, y1, snW, snH, 90, 90)
                        Call GdipDrawArc(hTmpGraphics, hPen, x1, y1, snW, snH, 180, 90)
                        Call GdipDrawArc(hTmpGraphics, hPen, x1, y1, snW, snH, 270, 90)
                    
                    Case [icoEllipseSolid]
                        Call GdipFillEllipse(hTmpGraphics, hBrush, x1, y1, snW, snH)
                        Call GdipDrawArc(hTmpGraphics, hPen, x1, y1, snW, snH, 0, 90)
                        Call GdipDrawArc(hTmpGraphics, hPen, x1, y1, snW, snH, 90, 90)
                        Call GdipDrawArc(hTmpGraphics, hPen, x1, y1, snW, snH, 180, 90)
                        Call GdipDrawArc(hTmpGraphics, hPen, x1, y1, snW, snH, 270, 90)
                    End Select
        End Select
          
        With uMatrix
            .m(0, 0) = 1
            .m(1, 1) = 1
            .m(2, 2) = 1
            .m(3, 3) = m_Alpha / 255
            .m(4, 4) = 1
        End With
        Call GdipCreateImageAttributes(hAttributes)
        Call GdipSetImageAttributesColorMatrix(hAttributes, [ColorAdjustTypeDefault], True, uMatrix, ByVal 0, [ColorMatrixFlagsDefault])
        Call GdipDrawImageRectRectI(hGraphics, hTmpBitmap, 0, 0, oDIB.Width, oDIB.Height, 0, 0, oDIB.Width, oDIB.Height, [UnitPixel], hAttributes, 0, 0)
        
        Call ARGBBitmapToGDIDIB32(oDIB, hBitmap)
        Call GdipDeletePen(hPen)
        Call GdipDeleteBrush(hBrush)
        Call GdipDeleteGraphics(hGraphics)
        Call GdipDeleteGraphics(hTmpGraphics)
        Call GdipDisposeImage(hBitmap)
        Call GdipDisposeImage(hTmpBitmap)
        Call GdipDisposeImageAttributes(hAttributes)

      Else
      
        lColorPen = pvColorRef(oDIB, Layer, [icoSampleA], 0)
        With uLogBrush
            .lbStyle = BS_SOLID
            .lbColor = lColorPen
            .lbHatch = 0
        End With
        If (m_Shape > [icoRectangleSolid]) Then
            hPen = ExtCreatePen(PS_GEOMETRIC + PS_ENDCAP_ROUND + PS_JOIN_ROUND, m_ShapeLineWidth, uLogBrush, 0, ByVal 0)
          Else
            hPen = ExtCreatePen(PS_GEOMETRIC + PS_ENDCAP_SQUARE + PS_JOIN_MITER, IIf(x1 = x2 And y1 = y2, 1, m_ShapeLineWidth), uLogBrush, 0, ByVal 0)
        End If
        hOldPen = SelectObject(oDIB.hDC, hPen)
        
        lColorBrush = pvColorRef(oDIB, Layer, [icoSampleA], 0)
        If (m_Shape = [icoRectangle] Or m_Shape = [icoEllipse]) Then
            With uLogBrush
                .lbStyle = BS_NULL
                .lbColor = 0
                .lbHatch = 0
            End With
            hBrush = CreateBrushIndirect(uLogBrush)
          Else
            hBrush = CreateSolidBrush(lColorBrush)
        End If
        hOldBrush = SelectObject(oDIB.hDC, hBrush)
                            
        Select Case True
        
            Case x1 = x2 And y1 = y2
            
                Select Case m_Shape
                    Case [icoRectangle], [icoRectangleSolid]
                        If (m_ShapeLineWidth = 1) Then
                            Call SetPixelV(oDIB.hDC, x2, y2, lColorPen)
                          Else
                            For y1 = y1 - m_ShapeLineWidth \ 2 To y2 + m_ShapeLineWidth \ 2
                                Call MoveToEx(oDIB.hDC, x1 - m_ShapeLineWidth \ 2, y1, ptFrom)
                                Call LineTo(oDIB.hDC, x2 + m_ShapeLineWidth \ 2, y1)
                            Next y1
                        End If
                    Case [icoEllipse], [icoEllipseSolid]
                        Call MoveToEx(oDIB.hDC, x1, y1, ptFrom)
                        Call LineTo(oDIB.hDC, x2, y2)
                        Call SetPixelV(oDIB.hDC, x2, y2, lColorPen)
                End Select
                
            Case x1 = x2 Xor y1 = y2
              
                Call MoveToEx(oDIB.hDC, x1, y1, ptFrom)
                Call LineTo(oDIB.hDC, x2, y2)
        
            Case Else
            
                Call pvCorrectShapeCoords(x1, y1, x2, y2)
                Select Case m_Shape
                    Case [icoRectangle], [icoRectangleSolid]
                        Call Rectangle(oDIB.hDC, x1, y1, x2, y2)
                    Case [icoEllipse], [icoEllipseSolid]
                        Call Ellipse(oDIB.hDC, x1, y1, x2, y2)
                End Select
        End Select
        
        Call SelectObject(oDIB.hDC, hOldPen)
        Call DeleteObject(hPen)
        Call SelectObject(oDIB.hDC, hOldBrush)
        Call DeleteObject(hBrush)
    End If
End Sub

Friend Sub MeasureString(oDIB As cDIB, _
           W As Single, H As Single, _
           ByVal Is32bpp As Boolean)

  Dim hGraphics As Long
  Dim hFontFam  As Long
  Dim hCurFont  As Long
  Dim rctFText  As RECTF
  Dim hFont     As Long
  Dim hOldFont  As Long
  Dim rctLText  As RECT2
        
    If (Is32bpp) Then
    
        Call GdipCreateFromHDC(oDIB.hDC, hGraphics)
        Call GdipSetTextRenderingHint(hGraphics, [TextRenderingHintSingleBitPerPixelGridFit] - 3 * m_Antialias)
      
        Call GdipCreateFontFamilyFromName(StrConv(m_oFont.Name, vbUnicode), 0, hFontFam)
        Call GdipCreateFont(hFontFam, m_oFont.Size, [FontStyleBold] * -m_oFont.Bold + [FontStyleItalic] * -m_oFont.Italic, [UnitPoint], hCurFont)
        
        Call GdipMeasureString(hGraphics, StrConv(m_Text, vbUnicode), -1, hCurFont, rctFText, 0, rctFText, 0, 0)
        With rctFText
            W = .W
            H = .H
        End With
        
        Call GdipDeleteFont(hCurFont)
        Call GdipDeleteFontFamily(hFontFam)
        Call GdipDeleteGraphics(hGraphics)
        
      Else
      
        hFont = CreateFontIndirect(pvOLEFontToLogFont(m_oFont, oDIB.hDC))
        hOldFont = SelectObject(oDIB.hDC, hFont)
        
        Call DrawText(oDIB.hDC, m_Text, -1, rctLText, DT_CALCRECT)
        With rctLText
            W = .x2 - .x1
            H = .y2 - .y1
        End With
        
        Call SelectObject(oDIB.hDC, hOldFont)
        Call DeleteObject(hFont)
    End If
End Sub

Friend Sub DrawString(oDIB As cDIB, _
           ByVal x As Single, ByVal y As Single, _
           ByVal W As Single, ByVal H As Single, _
           ByVal Layer As icoLayerCts)
           
  Dim hBitmap      As Long
  Dim hGraphics    As Long
  Dim lColorPen    As Long
  Dim lOldColorPen As Long
  Dim hBrush       As Long
  Dim hFontFam     As Long
  Dim hCurFont     As Long
  Dim rctFText     As RECTF
  Dim hFont        As Long
  Dim hOldFont     As Long
  Dim lOldBkMode   As Long
  Dim rctLText     As RECT2
        
    If (Layer = [icoXORAlpha]) Then
        
        Call ARGBBitmapFromGDIDIB32(oDIB, hBitmap)
        
        Call GdipGetImageGraphicsContext(hBitmap, hGraphics)
        Call GdipSetTextRenderingHint(hGraphics, [TextRenderingHintSingleBitPerPixelGridFit] - 3 * m_Antialias)
        
        Call GdipCreateFontFamilyFromName(StrConv(m_oFont.Name, vbUnicode), 0, hFontFam)
        Call GdipCreateFont(hFontFam, m_oFont.Size, [FontStyleBold] * -m_oFont.Bold + [FontStyleItalic] * -m_oFont.Italic, [UnitPoint], hCurFont)
        
        lColorPen = pvColorRef(oDIB, Layer, [icoSampleA], m_Alpha)
        Call GdipCreateSolidFill(lColorPen, hBrush)
        
        With rctFText
            .x = x: .W = W
            .y = y: .H = H
        End With
        Call GdipDrawString(hGraphics, StrConv(m_Text, vbUnicode), -1, hCurFont, rctFText, 0, hBrush)
        
        Call ARGBBitmapToGDIDIB32(oDIB, hBitmap)
        Call GdipDeletePen(hBrush)
        Call GdipDeleteFont(hCurFont)
        Call GdipDeleteFontFamily(hFontFam)
        Call GdipDeleteGraphics(hGraphics)
        Call GdipDisposeImage(hBitmap)
        
      Else
      
        hFont = CreateFontIndirect(pvOLEFontToLogFont(m_oFont, oDIB.hDC))
        hOldFont = SelectObject(oDIB.hDC, hFont)
        
        lOldBkMode = SetBkMode(oDIB.hDC, TRANSPARENT)
        
        lColorPen = pvColorRef(oDIB, Layer, [icoSampleA], 0)
        lOldColorPen = SetTextColor(oDIB.hDC, lColorPen)
        
        With rctLText
            .x1 = x: .x2 = x + W
            .y1 = y: .y2 = y + H
        End With
        Call DrawText(oDIB.hDC, m_Text, -1, rctLText, 0)
        
        Call SetTextColor(oDIB.hDC, lOldColorPen)
        Call SetBkMode(oDIB.hDC, lOldBkMode)
        Call SelectObject(oDIB.hDC, hOldFont)
        Call DeleteObject(hFont)
    End If
End Sub

Friend Sub ProcessColorScreen(oXORDIBCurr As cDIB, oXORDIBPrev As cDIB, oAlphaDIB As cDIB)

  Dim x As Long, y As Long
  Dim W As Long, H As Long

  Dim uSABitsA    As SAFEARRAY2D
  Dim uSABitsCurr As SAFEARRAY2D
  Dim uSABitsPrev As SAFEARRAY2D
  Dim aBitsA()    As Byte
  Dim aBitsCurr() As Byte
  Dim aBitsPrev() As Byte
  Dim lAlpha      As Long

    If (oXORDIBCurr.BPP = [32_bpp]) Then
    
        '-- Map DIBs
        Call pvMapDIB(oAlphaDIB, aBitsA(), uSABitsA)
        Call pvMapDIB(oXORDIBCurr, aBitsCurr(), uSABitsCurr)
        Call pvMapDIB(oXORDIBPrev, aBitsPrev(), uSABitsPrev)
    
        '-- Set alpha (substract alpha)
        W = 4 * oAlphaDIB.Width - 1
        H = 1 * oAlphaDIB.Height - 1
        For y = 0 To H
            For x = 0 To W Step 4
                If (aBitsA(x + 3, y)) Then
                    lAlpha = CLng(aBitsCurr(x + 3, y)) - aBitsA(x + 3, y): If (lAlpha < 0) Then lAlpha = 0
                    If (lAlpha) Then
                        aBitsCurr(x + 0, y) = aBitsPrev(x + 0, y)
                        aBitsCurr(x + 1, y) = aBitsPrev(x + 1, y)
                        aBitsCurr(x + 2, y) = aBitsPrev(x + 2, y)
                      Else
                        aBitsCurr(x + 0, y) = 0
                        aBitsCurr(x + 1, y) = 0
                        aBitsCurr(x + 2, y) = 0
                    End If
                    aBitsCurr(x + 3, y) = lAlpha
                  End If
            Next x
        Next y
    
        '-- Unmap DIBs
        Call pvUnmapDIB(aBitsA())
        Call pvUnmapDIB(aBitsCurr())
        Call pvUnmapDIB(aBitsPrev())
    End If
End Sub

Friend Sub FloodFill(oXORDIB As cDIB, oANDDIB As cDIB, ByVal x As Long, ByVal y As Long)

  Dim W As Long, xIn As Long, xInL As Long
  Dim H As Long, yIn As Long, yInL As Long
  
  Dim lClrXORFrom As Long
  Dim lClrANDFrom As Long
  Dim lClrXORTo   As Long
  Dim lClrANDTo   As Long
  Dim Rt As Byte, Gt As Byte, Bt As Byte, At As Byte
  Dim Rf As Byte, Gf As Byte, Bf As Byte, Af As Byte, HSLf As HSL
  Dim HSLc As HSL
  
  Dim hXORBitmap   As Long, hFillBitmap As Long
  Dim hXORGraphics As Long
  Dim oTmpXORDIB   As New cDIB
  Dim oFillDIB     As New cDIB
  
  Dim uSAXORBits  As SAFEARRAY2D
  Dim aXORBits()  As Byte
  Dim uSAFillBits As SAFEARRAY2D
  Dim aFillBits() As Byte
  
  Dim bProcessColorScreen As Boolean
  
  Dim rctBounds     As RECT2
  Dim nPoint        As Long
  Dim nPoints       As Long
  Dim nCheckPoint   As Long
  Dim nCheckRound   As Long
  Dim Points()      As POINTAPI
  Dim PointOffset   As POINTAPI
  Dim Point         As POINTAPI
  Dim bKeepChecking As Boolean
    
    Call SetRect(rctBounds, 0, 0, oXORDIB.Width, oXORDIB.Height)
    
    nPoint = 0
    nPoints = 0
    ReDim Points(nPoints) As POINTAPI
    
    Point.x = x
    Point.y = y
    Points(0) = Point
    
    x = Int(x)
    y = Int(y)
    W = oXORDIB.Width - 1
    H = oXORDIB.Height - 1
    ReDim bDone(W, H) As Boolean
                            
    Select Case oXORDIB.BPP
        
        Case [01_bpp], [04_bpp], [08_bpp], [24_bpp]
    
            lClrXORFrom = GetPixel(oXORDIB.hDC, x, y)
            lClrANDFrom = GetPixel(oANDDIB.hDC, x, y)
            
            lClrXORTo = pvColorRef(oXORDIB, [icoXORSolid], [icoSampleA], 0)
            lClrANDTo = pvColorRef(oANDDIB, [icoAND], [icoSampleA], 0)
            
            Call SetPixelV(oXORDIB.hDC, x, y, lClrXORTo)
            Call SetPixelV(oANDDIB.hDC, x, y, lClrANDTo)
            
            Do: bKeepChecking = False
        
                For nCheckPoint = nPoint To nPoints
        
                    For nCheckRound = 0 To 3
        
                        Select Case nCheckRound
                            Case 0: PointOffset.x = -1: PointOffset.y = 0
                            Case 1: PointOffset.x = 1:  PointOffset.y = 0
                            Case 2: PointOffset.x = 0:  PointOffset.y = -1
                            Case 3: PointOffset.x = 0:  PointOffset.y = 1
                        End Select
        
                        xIn = Points(nCheckPoint).x + PointOffset.x
                        yIn = Points(nCheckPoint).y + PointOffset.y
        
                        If (PtInRect(rctBounds, xIn, yIn)) Then
                            
                            If (bDone(xIn, yIn) = False) Then
                                bDone(xIn, yIn) = True
                            
                                If (GetPixel(oXORDIB.hDC, xIn, yIn) = lClrXORFrom And _
                                    GetPixel(oANDDIB.hDC, xIn, yIn) = lClrANDFrom) Then
                                    
                                    Call SetPixelV(oXORDIB.hDC, xIn, yIn, lClrXORTo)
                                    Call SetPixelV(oANDDIB.hDC, xIn, yIn, lClrANDTo)
                        
                                    nPoints = nPoints + 1
                                    ReDim Preserve Points(nPoints)
                                    Points(nPoints).x = xIn
                                    Points(nPoints).y = yIn
            
                                    bKeepChecking = True
                                End If
                            End If
                        End If
                    Next nCheckRound
                Next nCheckPoint

                nPoint = nCheckPoint

            Loop Until Not bKeepChecking
                            
        Case [32_bpp]
        
            bProcessColorScreen = (m_IsAScreen And Not m_SwapColors) Or _
                                  (m_IsBScreen And m_SwapColors)
            
            Call pvMapDIB(oXORDIB, aXORBits(), uSAXORBits)
            For yIn = 0 To H
                For xIn = 3 To 4 * W - 1 Step 4
                    If (aXORBits(xIn, yIn) = 0) Then
                        aXORBits(xIn - 1, yIn) = 0
                        aXORBits(xIn - 2, yIn) = 0
                        aXORBits(xIn - 3, yIn) = 0
                    End If
                Next xIn
            Next yIn
            
            Call oFillDIB.Create(oXORDIB.Width, oXORDIB.Height, [32_bpp])
            Call oFillDIB.Reset
            Call pvMapDIB(oFillDIB, aFillBits(), uSAFillBits)
            
            lClrXORTo = pvColorRef(oXORDIB, [icoXORAlpha], [icoSampleA], m_Alpha)
            At = m_Alpha
            Rt = (lClrXORTo And &HFF0000) \ 65536
            Gt = (lClrXORTo And &HFF00&) \ 256
            Bt = (lClrXORTo And &HFF&)
            
            Af = aXORBits(4 * x + 3, H - y)
            Rf = aXORBits(4 * x + 2, H - y)
            Gf = aXORBits(4 * x + 1, H - y)
            Bf = aXORBits(4 * x + 0, H - y)
            HSLf = pvRGBToHSL(Rf, Gf, Bf)
            
            aFillBits(4 * x + 3, H - y) = At
            aFillBits(4 * x + 2, H - y) = Rt
            aFillBits(4 * x + 1, H - y) = Gt
            aFillBits(4 * x + 0, H - y) = Bt
            
            Do: bKeepChecking = False
        
                For nCheckPoint = nPoint To nPoints
        
                    For nCheckRound = 0 To 3
        
                        Select Case nCheckRound
                            Case 0: PointOffset.x = -1: PointOffset.y = 0
                            Case 1: PointOffset.x = 1:  PointOffset.y = 0
                            Case 2: PointOffset.x = 0:  PointOffset.y = -1
                            Case 3: PointOffset.x = 0:  PointOffset.y = 1
                        End Select
        
                        xIn = Points(nCheckPoint).x + PointOffset.x
                        yIn = Points(nCheckPoint).y + PointOffset.y
        
                        If (PtInRect(rctBounds, xIn, yIn)) Then
                            
                            If (bDone(xIn, yIn) = False) Then
                                bDone(xIn, yIn) = True
                                
                                xInL = 4 * xIn
                                yInL = H - yIn
                                HSLc = pvRGBToHSL(aXORBits(xInL + 2, yInL), aXORBits(xInL + 1, yInL), aXORBits(xInL + 0, yInL))
                                
                                If (Abs(CLng(aXORBits(xInL + 3, yInL)) - Af) <= FillToleranceA And _
                                    Abs(HSLc.H - HSLf.H) <= FillToleranceH And _
                                    Abs(HSLc.S - HSLf.S) <= FillToleranceS And _
                                    Abs(HSLc.L - HSLf.L) <= FillToleranceL) Then
                                    
                                    aFillBits(xInL + 3, yInL) = At * (1 - (Abs(HSLc.H - HSLf.H) * Abs(HSLc.S - HSLf.S) * Abs(HSLc.L - HSLf.L)) / 3600000)
                                    aFillBits(xInL + 2, yInL) = Rt
                                    aFillBits(xInL + 1, yInL) = Gt
                                    aFillBits(xInL + 0, yInL) = Bt
                                    
                                    nPoints = nPoints + 1
                                    ReDim Preserve Points(nPoints)
                                    Points(nPoints).x = xIn
                                    Points(nPoints).y = yIn
                                    
                                    bKeepChecking = True
                                End If
                            End If
                        End If
                    Next nCheckRound
                Next nCheckPoint

                nPoint = nCheckPoint

            Loop Until Not bKeepChecking
    
            Call pvUnmapDIB(aXORBits())
            Call pvUnmapDIB(aFillBits())
            
            With oXORDIB
                
                If (bProcessColorScreen) Then
                    Call .CloneTo(oTmpXORDIB)
                End If
                
                Call ARGBBitmapFromGDIDIB32(oXORDIB, hXORBitmap)
                Call GdipGetImageGraphicsContext(hXORBitmap, hXORGraphics)
                Call ARGBBitmapFromGDIDIB32(oFillDIB, hFillBitmap)
                
                Call GdipDrawImageRectI(hXORGraphics, hFillBitmap, 0, 0, .Width, .Height)
                
                Call ARGBBitmapToGDIDIB32(oXORDIB, hXORBitmap)
                Call GdipDeleteGraphics(hXORGraphics)
                Call GdipDisposeImage(hXORBitmap)
                Call GdipDisposeImage(hFillBitmap)
                
                If (bProcessColorScreen) Then
                    Call Me.ProcessColorScreen(oXORDIB, oTmpXORDIB, oFillDIB)
                End If
            End With
    End Select
End Sub

Friend Sub FlipHorizontaly(oDIB As cDIB)

  Dim uSABits As SAFEARRAY2D
  Dim aBits() As Byte
  Dim aScan() As Byte
  Dim aIdx    As Byte
  
  Dim x As Long, lxMsk As Long
  Dim W As Long, lWMsk As Long
  Dim y As Long

    Call pvMapDIB(oDIB, aBits(), uSABits)

    With oDIB
  
        ReDim aScan(0 To .BytesPerScanline - 1)
        
        For y = 0 To .Height - 1
            
            x = 0
            W = .Width - 1
            
            Call CopyMemory(aScan(0), aBits(0, y), .BytesPerScanline)
            
            Do: Select Case .BPP
    
                    Case [01_bpp]
                        
                        lWMsk = m_Pow2(7 - W Mod 8)
                        aIdx = (aScan(W \ 8) And lWMsk) \ lWMsk
                        
                        lxMsk = m_Pow2(7 - x Mod 8)
                        If (aIdx = 1) Then
                            lxMsk = (aBits(x \ 8, y) And Not lxMsk) Or lxMsk
                          Else
                            lxMsk = (aBits(x \ 8, y) And Not lxMsk)
                        End If
                        aBits(x \ 8, y) = lxMsk And &HFF
    
                    Case [04_bpp]
                        
                        If ((W And 1) = 1) Then
                            aIdx = (aScan(W \ 2) And &HF&)
                          Else
                            aIdx = (aScan(W \ 2) And &HF0&) \ 16
                        End If
                        If ((x And 1) = 1) Then
                            aBits(x \ 2, y) = (aBits(x \ 2, y) And &HF0) Or aIdx
                          Else
                            aBits(x \ 2, y) = (aBits(x \ 2, y) And &HF) Or (aIdx * &H10)
                        End If
                        
                    Case [08_bpp]
                        
                        aBits(x, y) = aScan(W)
                    
                    Case [24_bpp]
                        
                        lWMsk = 3 * W
                        lxMsk = 3 * x
                        aBits(lxMsk + 0, y) = aScan(lWMsk + 0)
                        aBits(lxMsk + 1, y) = aScan(lWMsk + 1)
                        aBits(lxMsk + 2, y) = aScan(lWMsk + 2)
                    
                    Case [32_bpp]
                        
                        lWMsk = 4 * W
                        lxMsk = 4 * x
                        aBits(lxMsk + 0, y) = aScan(lWMsk + 0)
                        aBits(lxMsk + 1, y) = aScan(lWMsk + 1)
                        aBits(lxMsk + 2, y) = aScan(lWMsk + 2)
                        aBits(lxMsk + 3, y) = aScan(lWMsk + 3)
                End Select
                
                x = x + 1
                W = W - 1
            Loop Until W < 0
        Next y
    End With
    Call pvUnmapDIB(aBits())
End Sub

Friend Sub FlipVerticaly(oDIB As cDIB)
    
  Dim uSABits As SAFEARRAY2D
  Dim aBits() As Byte
  Dim aScan() As Byte, y As Long
  
    Call pvMapDIB(oDIB, aBits(), uSABits)

    With oDIB

        ReDim aScan(0 To .BytesPerScanline - 1)

        For y = 0 To .Height \ 2 - 1
            Call CopyMemory(aScan(0), aBits(0, y), .BytesPerScanline)
            Call CopyMemory(aBits(0, y), aBits(0, .Height - y - 1), .BytesPerScanline)
            Call CopyMemory(aBits(0, .Height - y - 1), aScan(0), .BytesPerScanline)
        Next y
    End With
    
    Call pvUnmapDIB(aBits())
End Sub

Friend Sub RotateLeft(oDIB As cDIB)

  Dim uSABits As SAFEARRAY2D
  Dim aBits() As Byte, aTmpBits() As Byte, aPal() As Byte
  Dim aIdx    As Byte
  
  Dim x As Long, W As Long, xR As Long
  Dim y As Long, H As Long
  Dim lMsk1 As Long
  Dim lMsk2 As Long
    
    Call pvMapDIB(oDIB, aBits(), uSABits): aTmpBits() = aBits()
    Call pvUnmapDIB(aBits())
    
    If (oDIB.BPP <= [08_bpp]) Then
        ReDim aPal(4 * 2 ^ oDIB.BPP - 1)
        Call oDIB.GetPalette(aPal())
    End If
    
    Call oDIB.Create(oDIB.Height, oDIB.Width, oDIB.BPP)
    Call pvMapDIB(oDIB, aBits(), uSABits)
    
    If (oDIB.BPP <= [08_bpp]) Then
        Call oDIB.SetPalette(aPal())
    End If
    
    With oDIB
    
        W = .Height - 1
        H = .Width - 1
        
        For y = 0 To H
            
            xR = H - y
            For x = 0 To W
                
                Select Case .BPP
    
                    Case [01_bpp]
                        
                        lMsk1 = m_Pow2(7 - x Mod 8)
                        aIdx = (aTmpBits(x \ 8, y) And lMsk1) \ lMsk1
                        
                        lMsk2 = m_Pow2(7 - xR Mod 8)
                        If (aIdx = 1) Then
                            lMsk2 = (aBits(xR \ 8, x) And Not lMsk2) Or lMsk2
                          Else
                            lMsk2 = (aBits(xR \ 8, x) And Not lMsk2)
                        End If
                        aBits(xR \ 8, x) = lMsk2 And &HFF
    
                    Case [04_bpp]
                        
                        If ((x And 1) = 1) Then
                            aIdx = (aTmpBits(x \ 2, y) And &HF&)
                          Else
                            aIdx = (aTmpBits(x \ 2, y) And &HF0&) \ 16
                        End If
                        If ((xR And 1) = 1) Then
                            aBits(xR \ 2, x) = (aBits(xR \ 2, x) And &HF0) Or aIdx
                          Else
                            aBits(xR \ 2, x) = (aBits(xR \ 2, x) And &HF) Or (aIdx * &H10)
                        End If
                        
                    Case [08_bpp]
                        
                        aBits(xR, x) = aTmpBits(x, y)
                    
                    Case [24_bpp]
                        
                        lMsk1 = 3 * x
                        lMsk2 = 3 * xR
                        aBits(lMsk2 + 0, x) = aTmpBits(lMsk1 + 0, y)
                        aBits(lMsk2 + 1, x) = aTmpBits(lMsk1 + 1, y)
                        aBits(lMsk2 + 2, x) = aTmpBits(lMsk1 + 2, y)
                    
                    Case [32_bpp]
                        
                        lMsk1 = 4 * x
                        lMsk2 = 4 * xR
                        aBits(lMsk2 + 0, x) = aTmpBits(lMsk1 + 0, y)
                        aBits(lMsk2 + 1, x) = aTmpBits(lMsk1 + 1, y)
                        aBits(lMsk2 + 2, x) = aTmpBits(lMsk1 + 2, y)
                        aBits(lMsk2 + 3, x) = aTmpBits(lMsk1 + 3, y)
                End Select
            Next x
        Next y
    End With

    Call pvUnmapDIB(aBits())
End Sub

Friend Sub RotateRight(oDIB As cDIB)

  Dim uSABits As SAFEARRAY2D
  Dim aBits() As Byte, aTmpBits() As Byte, aPal() As Byte
  Dim aIdx    As Byte
  
  Dim x As Long, W As Long
  Dim y As Long, H As Long, yR As Long
  Dim lMsk1 As Long
  Dim lMsk2 As Long
    
    Call pvMapDIB(oDIB, aBits(), uSABits): aTmpBits() = aBits()
    Call pvUnmapDIB(aBits())
    
    If (oDIB.BPP <= [08_bpp]) Then
        ReDim aPal(4 * 2 ^ oDIB.BPP - 1)
        Call oDIB.GetPalette(aPal())
    End If
    
    Call oDIB.Create(oDIB.Height, oDIB.Width, oDIB.BPP)
    Call pvMapDIB(oDIB, aBits(), uSABits)
    
    If (oDIB.BPP <= [08_bpp]) Then
        Call oDIB.SetPalette(aPal())
    End If
    
    With oDIB
    
        W = .Height - 1
        H = .Width - 1
    
        For x = 0 To W
            
            yR = W - x
            For y = 0 To H
            
                Select Case .BPP
    
                    Case [01_bpp]
                        
                        lMsk1 = m_Pow2(7 - x Mod 8)
                        aIdx = (aTmpBits(x \ 8, y) And lMsk1) \ lMsk1
                        
                        lMsk2 = m_Pow2(7 - y Mod 8)
                        If (aIdx = 1) Then
                            lMsk2 = (aBits(y \ 8, yR) And Not lMsk2) Or lMsk2
                          Else
                            lMsk2 = (aBits(y \ 8, yR) And Not lMsk2)
                        End If
                        aBits(y \ 8, yR) = lMsk2 And &HFF
    
                    Case [04_bpp]
                        
                        If ((x And 1) = 1) Then
                            aIdx = (aTmpBits(x \ 2, y) And &HF&)
                          Else
                            aIdx = (aTmpBits(x \ 2, y) And &HF0&) \ 16
                        End If
                        If ((y And 1) = 1) Then
                            aBits(y \ 2, yR) = (aBits(y \ 2, yR) And &HF0) Or aIdx
                          Else
                            aBits(y \ 2, yR) = (aBits(y \ 2, yR) And &HF) Or (aIdx * &H10)
                        End If
                        
                    Case [08_bpp]
                        
                        aBits(y, yR) = aTmpBits(x, y)
                    
                    Case [24_bpp]
                        
                        lMsk1 = 3 * x
                        lMsk2 = 3 * y
                        aBits(lMsk2 + 0, yR) = aTmpBits(lMsk1 + 0, y)
                        aBits(lMsk2 + 1, yR) = aTmpBits(lMsk1 + 1, y)
                        aBits(lMsk2 + 2, yR) = aTmpBits(lMsk1 + 2, y)
                    
                    Case [32_bpp]
                        
                        lMsk1 = 4 * x
                        lMsk2 = 4 * y
                        aBits(lMsk2 + 0, yR) = aTmpBits(lMsk1 + 0, y)
                        aBits(lMsk2 + 1, yR) = aTmpBits(lMsk1 + 1, y)
                        aBits(lMsk2 + 2, yR) = aTmpBits(lMsk1 + 2, y)
                        aBits(lMsk2 + 3, yR) = aTmpBits(lMsk1 + 3, y)
                End Select
            Next y
        Next x
    End With

    Call pvUnmapDIB(aBits())
End Sub

Friend Sub CropIcon(oCropRect As cRect, _
                    oXORDIB As cDIB, oANDDIB As cDIB)
      
  Dim oXORTmpDIB As New cDIB
  Dim oANDTmpDIB As New cDIB
  Dim aPal()     As Byte
  Dim oldAlpha   As Byte
  Dim oDstRect   As cRect
  
    Call oXORDIB.CloneTo(oXORTmpDIB)
    Call oANDDIB.CloneTo(oANDTmpDIB)
    
    With oCropRect
        Call .CloneTo(oDstRect)
        Call oDstRect.Offset(-.x1, -.y1)
    End With
    With oDstRect
        Call oXORDIB.Create(.x2, .y2, oXORDIB.BPP)
        Call oANDDIB.Create(.x2, .y2, [01_bpp])
    End With
    
    If (oXORDIB.BPP <= [08_bpp]) Then
        ReDim aPal(4 * m_Pow2(oXORTmpDIB.BPP) - 1)
        Call oXORTmpDIB.GetPalette(aPal())
        Call oXORDIB.SetPalette(aPal())
    End If
    ReDim aPal(7)
    Call oANDTmpDIB.GetPalette(aPal())
    Call oANDDIB.SetPalette(aPal())
    
    Call oXORDIB.Cls(&H0)
    Call oANDDIB.Cls(&HFFFFFF)
    
    oldAlpha = m_Alpha: m_Alpha = 255
    Call Me.MergeIcon(oDstRect, oCropRect, oXORDIB, oANDDIB, oXORTmpDIB, oANDTmpDIB)
    m_Alpha = oldAlpha
End Sub

Friend Sub MergeIcon(oDstRect As cRect, oSrcRect As cRect, _
                     oXORDstDIB As cDIB, oANDDstDIB As cDIB, _
                     oXORSrcDIB As cDIB, oANDSrcDIB As cDIB)
  
  Dim hDstBitmap    As Long, hSrcBitmap   As Long
  Dim hDstGraphics  As Long, hSrcGraphics As Long
  Dim hAttributes   As Long
  Dim uMatrix       As COLORMATRIX
                      
  Dim uSAXORDst     As SAFEARRAY2D
  Dim uSAANDDst     As SAFEARRAY2D
  Dim aXORDstBits() As Byte
  Dim aANDDstBits() As Byte
  Dim uSAXORSrc     As SAFEARRAY2D
  Dim uSAANDSrc     As SAFEARRAY2D
  Dim aXORSrcBits() As Byte
  Dim aANDSrcBits() As Byte
  Dim rctBounds     As RECT2
  
  Dim x As Long, xD As Long, xS As Long, xOff As Long
  Dim y As Long, yD As Long, yS As Long, yOff As Long, HD As Long, HS As Long
  
  Dim lMsk As Long
  Dim aIdx As Byte
       
    If (oXORDstDIB.BPP = [32_bpp]) Then
        
        Call ARGBBitmapFromGDIDIB32(oXORDstDIB, hDstBitmap)
        Call GdipGetImageGraphicsContext(hDstBitmap, hDstGraphics)
        Call ARGBBitmapFromGDIDIB32(oXORSrcDIB, hSrcBitmap)
        Call GdipGetImageGraphicsContext(hSrcBitmap, hSrcGraphics)
        
        With uMatrix
            .m(0, 0) = 1
            .m(1, 1) = 1
            .m(2, 2) = 1
            .m(3, 3) = m_Alpha / 255
            .m(4, 4) = 1
        End With
        Call GdipCreateImageAttributes(hAttributes)
        Call GdipSetImageAttributesColorMatrix(hAttributes, [ColorAdjustTypeDefault], True, uMatrix, ByVal 0, [ColorMatrixFlagsDefault])
        Call GdipDrawImageRectRectI(hDstGraphics, hSrcBitmap, oDstRect.x1, oDstRect.y1, oDstRect.x2 - oDstRect.x1, oDstRect.y2 - oDstRect.y1, oSrcRect.x1, oSrcRect.y1, oSrcRect.x2 - oSrcRect.x1, oSrcRect.y2 - oSrcRect.y1, [UnitPixel], hAttributes, 0, 0)
        
        Call ARGBBitmapToGDIDIB32(oXORDstDIB, hDstBitmap)
        Call GdipDeleteGraphics(hDstGraphics)
        Call GdipDisposeImage(hDstBitmap)
        Call GdipDeleteGraphics(hSrcGraphics)
        Call GdipDisposeImage(hSrcBitmap)
        Call GdipDisposeImageAttributes(hAttributes)
      
      Else
      
        Call pvMapDIB(oXORDstDIB, aXORDstBits(), uSAXORDst)
        Call pvMapDIB(oANDDstDIB, aANDDstBits(), uSAANDDst)
        Call pvMapDIB(oXORSrcDIB, aXORSrcBits(), uSAXORSrc)
        Call pvMapDIB(oANDSrcDIB, aANDSrcBits(), uSAANDSrc)

        Call SetRect(rctBounds, 0, 0, oXORDstDIB.Width, oXORDstDIB.Height)
        xOff = oSrcRect.x1 - oDstRect.x1
        yOff = oSrcRect.y1 - oDstRect.y1
        HD = oXORDstDIB.Height - 1
        HS = oXORSrcDIB.Height - 1
        
        For y = oDstRect.y1 To oDstRect.y2 - 1
            For x = oDstRect.x1 To oDstRect.x2 - 1
                
                xD = x
                yD = HD - y
                xS = x + xOff
                yS = HS - (y + yOff)
                
                If (PtInRect(rctBounds, xD, yD)) Then
                
                    lMsk = m_Pow2(7 - xS Mod 8)
                    aIdx = (aANDSrcBits(xS \ 8, yS) And lMsk) \ lMsk
                    
                    '-- Not screen (black)
                    If (aIdx = 0) Then
                        
                        '-- Draw onto AND DIB
                        lMsk = m_Pow2(7 - xD Mod 8)
                        lMsk = (aANDDstBits(xD \ 8, yD) And Not lMsk)
                        aANDDstBits(xD \ 8, yD) = lMsk And &HFF&
                        
                        '-- Draw onto XOR DIB
                        Select Case oXORDstDIB.BPP
                        
                            Case [01_bpp]
                                
                                lMsk = m_Pow2(7 - xS Mod 8)
                                aIdx = (aXORSrcBits(xS \ 8, yS) And lMsk) \ lMsk
                                
                                lMsk = m_Pow2(7 - xD Mod 8)
                                If (aIdx = 1) Then
                                    lMsk = (aXORDstBits(xD \ 8, yD) And Not lMsk) Or lMsk
                                  Else
                                    lMsk = (aXORDstBits(xD \ 8, yD) And Not lMsk)
                                End If
                                aXORDstBits(xD \ 8, yD) = lMsk And &HFF
    
                            Case [04_bpp]
                                
                                If ((xS And 1) = 1) Then
                                    aIdx = (aXORSrcBits(xS \ 2, yS) And &HF&)
                                  Else
                                    aIdx = (aXORSrcBits(xS \ 2, yS) And &HF0&) \ 16
                                End If
                                If ((xD And 1) = 1) Then
                                    aXORDstBits(xD \ 2, yD) = (aXORDstBits(xD \ 2, yD) And &HF0) Or aIdx
                                  Else
                                    aXORDstBits(xD \ 2, yD) = (aXORDstBits(xD \ 2, yD) And &HF) Or (aIdx * &H10)
                                End If
                                
                            Case [08_bpp]
                                
                                aXORDstBits(xD, yD) = aXORSrcBits(xD + xOff, yS)
                            
                            Case [24_bpp]
                                
                                aXORDstBits(3 * xD + 0, yD) = aXORSrcBits(3 * xS + 0, yS)
                                aXORDstBits(3 * xD + 1, yD) = aXORSrcBits(3 * xS + 1, yS)
                                aXORDstBits(3 * xD + 2, yD) = aXORSrcBits(3 * xS + 2, yS)
                        End Select
                    End If
                End If
            Next x
        Next y
        
        Call pvUnmapDIB(aXORSrcBits())
        Call pvUnmapDIB(aANDSrcBits())
        Call pvUnmapDIB(aXORDstBits())
        Call pvUnmapDIB(aANDDstBits())
    End If
End Sub

Friend Sub ClearIconRect(oClearRect As cRect, _
                         oXORDIB As cDIB, oANDDIB As cDIB)
    
  Dim rctClear As RECT2
  Dim hBrush   As Long
    
    With oClearRect
        Call SetRect(rctClear, .x1, .y1, .x2, .y2)
    End With
   
    hBrush = CreateSolidBrush(&H0)
    Call FillRect(oXORDIB.hDC, rctClear, hBrush)
    Call DeleteObject(hBrush)
    
    hBrush = CreateSolidBrush(&HFFFFFF)
    Call FillRect(oANDDIB.hDC, rctClear, hBrush)
    Call DeleteObject(hBrush)
End Sub

'//

Friend Sub ImportIconFromPrivateClipboard(oXORDIB As cDIB, oANDDIB As cDIB, oXORClipDIB As cDIB, oANDClipDIB As cDIB, oDstRect As cRect)
    
  Dim oDIB32 As New cDIB
  Dim aPal() As Byte
  Dim bfx As Long, bfy As Long
  Dim bfW As Long, bfH As Long
  
    '-- Store temp. palette [?]
    If (oXORDIB.BPP <= [08_bpp]) Then
        ReDim aPal(0 To 4 * 2 ^ oXORDIB.BPP - 1): Call oXORDIB.GetPalette(aPal())
    End If
    
    '-- Dest. rect.
    With oDstRect
        If (.IsEmpty) Then
            Call oXORDIB.GetBestFitInfo(oXORClipDIB.Width, oXORClipDIB.Height, oXORDIB.Width, oXORDIB.Height, bfx, bfy, bfW, bfH)
            Call .SetCoords(bfx, bfy, bfx + bfW, bfy + bfH)
          Else
            Call .GetCoords(bfx, bfy, bfW, bfH)
            bfW = bfW - bfx
            bfH = bfH - bfy
        End If
    End With
    
    '-- Create DIBs
    Call oXORDIB.Create(bfW, bfH, oXORDIB.BPP)
    If (oXORDIB.BPP <= [08_bpp]) Then
        Call oXORDIB.SetPalette(aPal())
    End If
    Call oANDDIB.Create(bfW, bfH, [01_bpp])
    ReDim aPal(7): Call FillMemory(aPal(4), 3, &HFF): Call oANDDIB.SetPalette(aPal())
        
    '-- Translate...
    Select Case True
        
        Case oXORDIB.BPP = [32_bpp]
        
            Call oXORDIB.Reset
            
            Select Case True
            
                Case oXORClipDIB.BPP = [32_bpp]
                    
                    If (m_Antialias) Then
                        Call pvStretchARGB(oXORClipDIB, oXORDIB)
                      Else
                        Call oXORClipDIB.Stretch(oXORDIB.hDC, 0, 0, oXORDIB.Width, oXORDIB.Height)
                    End If
        
                Case Else
            
                    If (m_Antialias) Then
                        With oDIB32
                            Call .Create(oXORClipDIB.Width, oXORClipDIB.Height, [32_bpp])
                            Call oXORClipDIB.Stretch(.hDC, 0, 0, .Width, .Height)
                        End With
                        Call ProcessAlphaIconFromMaskDIB(oDIB32, oANDClipDIB)
                        Call pvStretchARGB(oDIB32, oXORDIB)
                      Else
                        Call oXORClipDIB.Stretch(oXORDIB.hDC, 0, 0, oXORDIB.Width, oXORDIB.Height)
                        Call ProcessAlphaIconFromMaskDIB(oXORDIB, oANDClipDIB)
                    End If
                End Select
            
        Case Else
        
            Call oXORDIB.Cls(&H0)
            Call oANDDIB.Cls(&H0)
        
            Select Case True
            
                Case oXORClipDIB.BPP = [32_bpp]
            
                    Call oXORClipDIB.Stretch(oXORDIB.hDC, 0, 0, oXORDIB.Width, oXORDIB.Height)
                    Call ProcessSolidIconFromAlphaDIB(oXORDIB, oANDDIB, oXORClipDIB)
                
                Case Else
                    Call oXORClipDIB.Stretch(oXORDIB.hDC, 0, 0, oXORDIB.Width, oXORDIB.Height)
                    Call oANDClipDIB.Stretch(oANDDIB.hDC, 0, 0, oANDDIB.Width, oANDDIB.Height)
            End Select
   End Select
End Sub

Friend Sub ImportIconFromWindowsClipboard(oXORDIB As cDIB, oANDDIB As cDIB, oDstRect As cRect)
    
  Dim oClipDIB As New cDIB
  Dim oPal     As New cPalette
  Dim aPal()   As Byte
  Dim bfx As Long, bfy As Long
  Dim bfW As Long, bfH As Long
  
  Dim oDIB32     As New cDIB
  Dim uSA32      As SAFEARRAY2D
  Dim uSAXOR     As SAFEARRAY2D
  Dim uSAAND     As SAFEARRAY2D
  Dim a32Bits()  As Byte
  Dim aXORBits() As Byte
  Dim aANDBits() As Byte
  Dim aIdxBlack  As Byte
  Dim R          As Byte
  Dim G          As Byte
  Dim B          As Byte
  
  Dim x As Long, xIn As Long, W As Long
  Dim y As Long, H As Long
    
    '-- Import/create from Clipboard
    Call oClipDIB.CreateFromStdPicture(Clipboard.GetData(vbCFDIB), Force32bpp:=True)
    
    '-- Store temp. palette / Get nearest index (black) [?]
    If (oXORDIB.BPP <= [08_bpp]) Then
        ReDim aPal(0 To 4 * 2 ^ oXORDIB.BPP - 1): Call oXORDIB.GetPalette(aPal())
        Call oPal.Initialize(2 ^ oXORDIB.BPP)
        Call CopyMemory(ByVal oPal.lpPalette, aPal(0), 4 * oPal.Entries)
        Call oPal.BuildLogicalPalette
        Call oPal.ClosestIndex(0, 0, 0, aIdxBlack)
    End If
    
    '-- Dest. rect.
    With oDstRect
        If (.IsEmpty) Then
            Call oXORDIB.GetBestFitInfo(oClipDIB.Width, oClipDIB.Height, oXORDIB.Width, oXORDIB.Height, bfx, bfy, bfW, bfH)
            Call .SetCoords(bfx, bfy, bfx + bfW, bfy + bfH)
          Else
            Call .GetCoords(bfx, bfy, bfW, bfH)
            bfW = bfW - bfx
            bfH = bfH - bfy
        End If
    End With
    
    '-- Create DIBs
    Call oXORDIB.Create(bfW, bfH, oXORDIB.BPP)
    If (oXORDIB.BPP <= [08_bpp]) Then
        Call oXORDIB.SetPalette(aPal())
    End If
    Call oANDDIB.Create(bfW, bfH, [01_bpp])
    ReDim aPal(7): Call FillMemory(aPal(4), 3, &HFF): Call oANDDIB.SetPalette(aPal())
        
    '-- Translate...
    If (oXORDIB.BPP = [32_bpp] And m_Antialias) Then
        
        '-- Make opaque 32bpp (source)
        If (m_Import32bppOpaque) Then
            Screen.MousePointer = vbHourglass
            Call MakeOpaqueDIB32(oClipDIB)
            Screen.MousePointer = vbDefault
        End If
        
        '-- Interpolated
        Call oXORDIB.Reset
        Call pvStretchARGB(oClipDIB, oXORDIB)
      
      Else
        
        '-- Not interpolated
        If (oXORDIB.BPP = [32_bpp]) Then
            Call oXORDIB.Reset
          Else
            Call oXORDIB.Cls(&H0)
            Call oANDDIB.Cls(&H0)
        End If
        Call oClipDIB.Stretch(oXORDIB.hDC, 0, 0, oXORDIB.Width, oXORDIB.Height)
        
        '-- Make opaque 32bpp (target)
        If (oXORDIB.BPP = [32_bpp] And m_Import32bppOpaque) Then
            Call MakeOpaqueDIB32(oXORDIB)
        End If
    End If
    
    '-- Mask color B
    If (m_MaskImport) Then
        
        Call oDIB32.Create(oXORDIB.Width, oXORDIB.Height, [32_bpp])
        Call oClipDIB.Stretch(oDIB32.hDC, 0, 0, oXORDIB.Width, oXORDIB.Height)

        Call pvMapDIB(oDIB32, a32Bits(), uSA32)
        Call pvMapDIB(oXORDIB, aXORBits(), uSAXOR)
        Call pvMapDIB(oANDDIB, aANDBits(), uSAAND)
        
        W = oXORDIB.Width - 1
        H = oXORDIB.Height - 1
        
        R = (m_ColorLngB And &HFF0000) \ 65536
        G = (m_ColorLngB And &HFF00&) \ 256
        B = (m_ColorLngB And &HFF&)
      
        For y = 0 To H
            For x = 0 To W
                
                xIn = 4 * x
                If (a32Bits(xIn + 2, y) = R And a32Bits(xIn + 1, y) = G And a32Bits(xIn + 0, y) = B) Then
                
                    Select Case oXORDIB.BPP
                    
                        Case [01_bpp]
                            
                            xIn = x \ 8
                            aXORBits(xIn, y) = aXORBits(xIn, y) And Not m_Pow2(7 - (x Mod 8))
    
                            aANDBits(xIn, y) = aANDBits(xIn, y) Or m_Pow2(7 - (x Mod 8))
                        
                        Case [04_bpp]
                            
                            xIn = x \ 2
                            If ((x And 1) = 1) Then
                                aXORBits(xIn, y) = (aXORBits(xIn, y) And &HF0) Or aIdxBlack
                              Else
                                aXORBits(xIn, y) = (aXORBits(xIn, y) And &HF) Or (aIdxBlack * &H10)
                            End If
                            
                            xIn = x \ 8
                            aANDBits(xIn, y) = aANDBits(xIn, y) Or m_Pow2(7 - (x Mod 8))
                            
                        Case [08_bpp]
                            
                            aXORBits(x, y) = aIdxBlack
                            
                            xIn = x \ 8
                            aANDBits(xIn, y) = aANDBits(xIn, y) Or m_Pow2(7 - (x Mod 8))
                        
                        Case [24_bpp]
                            
                            xIn = 3 * x
                            aXORBits(xIn + 0, y) = &H0
                            aXORBits(xIn + 1, y) = &H0
                            aXORBits(xIn + 2, y) = &H0
                            
                            xIn = x \ 8
                            aANDBits(xIn, y) = aANDBits(xIn, y) Or m_Pow2(7 - (x Mod 8))
                            
                        Case [32_bpp]
                            
                            aXORBits(4 * x + 3, y) = &H0
                    End Select
                End If
            Next x
        Next y
        
        Call pvUnmapDIB(a32Bits())
        Call pvUnmapDIB(aXORBits())
        Call pvUnmapDIB(aANDBits())
    End If
End Sub

Friend Sub ExportIconToWindowsClipboard(oXORDIB As cDIB, oANDDIB As cDIB)
  
  Dim oXORBUffer As New cDIB
  Dim uSAXOR     As SAFEARRAY2D
  Dim uSAAND     As SAFEARRAY2D
  Dim aXORBits() As Byte
  Dim aANDBits() As Byte
  Dim lMsk       As Long
  Dim aIdx       As Long
  Dim R          As Byte
  Dim G          As Byte
  Dim B          As Byte
  
  Dim x As Long, xIn As Long, W As Long
  Dim y As Long, H As Long
    
    If (Not m_MaskExport Or oXORDIB.BPP = [32_bpp]) Then
    
        '-- Copy current XOR DIB
        Call oXORDIB.CopyToClipboard
    
      Else
        '-- Mask color B
        
        Call oXORDIB.CloneTo(oXORBUffer)
        Call pvMapDIB(oXORBUffer, aXORBits(), uSAXOR)
        Call pvMapDIB(oANDDIB, aANDBits(), uSAAND)
        
        W = oXORDIB.Width - 1
        H = oXORDIB.Height - 1
        
        R = (m_ColorLngB And &HFF0000) \ 65536
        G = (m_ColorLngB And &HFF00&) \ 256
        B = (m_ColorLngB And &HFF&)
      
        For y = 0 To H
            For x = 0 To W
                                
                lMsk = m_Pow2(7 - x Mod 8)
                aIdx = (aANDBits(x \ 8, y) And lMsk) \ lMsk
                
                '-- Screen (white)
                If (aIdx = 1) Then
                    
                    '-- Change XOR mask color
                    Select Case oXORDIB.BPP
                    
                        Case [01_bpp]

                            xIn = x \ 8
                            If (m_ColorIdxB = 0) Then
                                aXORBits(xIn, y) = aXORBits(xIn, y) And Not m_Pow2(7 - (x Mod 8))
                              Else
                                aXORBits(xIn, y) = aXORBits(xIn, y) Or m_Pow2(7 - (x Mod 8))
                            End If
                        
                        Case [04_bpp]
                            
                            xIn = x \ 2
                            If ((x And 1) = 1) Then
                                aXORBits(xIn, y) = (aXORBits(xIn, y) And &HF0) Or m_ColorIdxB
                              Else
                                aXORBits(xIn, y) = (aXORBits(xIn, y) And &HF) Or (m_ColorIdxB * &H10)
                            End If
                            
                        Case [08_bpp]
                            
                            aXORBits(x, y) = m_ColorIdxB
                        
                        Case [24_bpp]
                            
                            xIn = 3 * x
                            aXORBits(xIn + 0, y) = R
                            aXORBits(xIn + 1, y) = G
                            aXORBits(xIn + 2, y) = B
                    End Select
                End If
            Next x
        Next y
        
        Call pvUnmapDIB(aXORBits())
        Call pvUnmapDIB(aANDBits())
        
        Call oXORBUffer.CopyToClipboard
    End If
End Sub

Friend Sub ProcessSolidIconFromAlphaDIB(oXORDIB As cDIB, oANDDIB As cDIB, oAlphaDIB As cDIB)
    
  Dim oAlphaMask   As New cDIB
  Dim oPal         As New cPalette
  Dim aPal()       As Byte
  Dim aBlackIdx    As Byte
  Dim uSAXOR       As SAFEARRAY2D
  Dim uSAAND       As SAFEARRAY2D
  Dim uSAAlpha     As SAFEARRAY2D
  Dim aXORBits()   As Byte
  Dim aANDBits()   As Byte
  Dim aAlphaBits() As Byte
  
  Dim x As Long, W As Long
  Dim y As Long, H As Long
    
    Call oAlphaMask.Create(oXORDIB.Width, oXORDIB.Height, [32_bpp])
    Call oAlphaMask.Reset
    Call oAlphaDIB.Stretch(oAlphaMask.hDC, 0, 0, oAlphaMask.Width, oAlphaMask.Height)
    
    Call pvMapDIB(oXORDIB, aXORBits(), uSAXOR)
    Call pvMapDIB(oANDDIB, aANDBits(), uSAAND)
    Call pvMapDIB(oAlphaMask, aAlphaBits(), uSAAlpha)
        
    W = oAlphaMask.Width - 1
    H = oAlphaMask.Height - 1
    
    If (oXORDIB.BPP <= [08_bpp]) Then
    
        Call oPal.Initialize(2 ^ oXORDIB.BPP): ReDim aPal(4 * 2 ^ oXORDIB.BPP - 1)
        Call oXORDIB.GetPalette(aPal())
        Call CopyMemory(ByVal oPal.lpPalette, aPal(0), 4 * oPal.Entries)
        Call oPal.BuildLogicalPalette
        Call oPal.ClosestIndex(0, 0, 0, aBlackIdx)
    End If
    
    For y = 0 To H
        For x = 0 To W
        
            If (aAlphaBits(4 * x + 3, y) < 128) Then
                
                '-- XOR -> black index (closest)
                Select Case oXORDIB.BPP
                
                    Case [01_bpp] ' 0
                        
                        aXORBits(x \ 8, y) = aXORBits(x \ 8, y) And Not m_Pow2(7 - (x Mod 8))
                    
                    Case [04_bpp] ' Closest (Not always EGA. Should be.)
                        
                        If (x Mod 2 = 1) Then
                            aXORBits(x \ 2, y) = (aXORBits(x \ 2, y) And &HF0) Or aBlackIdx
                          Else
                            aXORBits(x \ 2, y) = (aXORBits(x \ 2, y) And &HF) Or (aBlackIdx * &H10)
                        End If
                    
                    Case [08_bpp] ' Closest
                        
                        aXORBits(x, y) = aBlackIdx
                        
                    Case [24_bpp] ' Black RGB
                        
                        aXORBits(3 * x + 0, y) = &H0
                        aXORBits(3 * x + 1, y) = &H0
                        aXORBits(3 * x + 2, y) = &H0
                End Select
                
                '-- AND -> white index (1)
                aANDBits(x \ 8, y) = aANDBits(x \ 8, y) Or m_Pow2(7 - (x Mod 8))
            End If
        Next x
    Next y
        
    Call pvUnmapDIB(aXORBits())
    Call pvUnmapDIB(aANDBits())
    Call pvUnmapDIB(aAlphaBits())
End Sub

Friend Sub ProcessAlphaIconFromMaskDIB(oXORDIB As cDIB, oMaskDIB As cDIB)
    
  Dim oMask       As New cDIB
  Dim aPal()      As Byte
  Dim uSAXOR      As SAFEARRAY2D
  Dim uSAMask     As SAFEARRAY2D
  Dim aXORBits()  As Byte
  Dim aMaskBits() As Byte
  Dim lMsk        As Long
  
  Dim x As Long, W As Long
  Dim y As Long, H As Long
    
    Call oMask.Create(oXORDIB.Width, oXORDIB.Height, [01_bpp])
    ReDim aPal(7): Call FillMemory(aPal(4), 3, &HFF)
    Call oMask.SetPalette(aPal())
    Call oMask.Reset
    Call oMaskDIB.Stretch(oMask.hDC, 0, 0, oMask.Width, oMask.Height)
    
    Call pvMapDIB(oXORDIB, aXORBits(), uSAXOR)
    Call pvMapDIB(oMask, aMaskBits(), uSAMask)
        
    W = oMask.Width - 1
    H = oMask.Height - 1
    
    For y = 0 To H
        For x = 0 To W
            
            lMsk = m_Pow2(7 - x Mod 8)
            If ((aMaskBits(x \ 8, y) And lMsk) \ lMsk = 0) Then
                aXORBits(4 * x + 3, y) = &HFF
              Else
                aXORBits(4 * x + 3, y) = &H0
            End If
        Next x
    Next y
        
    Call pvUnmapDIB(aXORBits())
    Call pvUnmapDIB(aMaskBits())
End Sub

Friend Sub MakeOpaqueDIB32(oDIB32 As cDIB)

  Dim uSA       As SAFEARRAY2D
  Dim aBits()   As Byte
  
  Dim x As Long, W As Long
  Dim y As Long, H As Long
    
    Call pvMapDIB(oDIB32, aBits(), uSA)

    W = oDIB32.Width - 1
    H = oDIB32.Height - 1

    For y = 0 To H
        For x = 3 To 4 * W + 3 Step 4
            aBits(x, y) = &HFF
        Next x
    Next y

    Call pvUnmapDIB(aBits())
End Sub

Friend Sub RemaskARGBIcon(oXORDIB As cDIB, oANDDIB As cDIB)
    
  Dim uSAXOR     As SAFEARRAY2D
  Dim uSAAND     As SAFEARRAY2D
  Dim aXORBits() As Byte
  Dim aANDBits() As Byte
  
  Dim x As Long, xIn As Long, W As Long
  Dim y As Long, H As Long
    
    If (oXORDIB.BPP = [32_bpp]) Then
    
        Call pvMapDIB(oXORDIB, aXORBits(), uSAXOR)
        Call pvMapDIB(oANDDIB, aANDBits(), uSAAND)
        
        W = oXORDIB.Width - 1
        H = oXORDIB.Height - 1
        
        For y = 0 To H
            For x = 0 To W
                
                xIn = 4 * x
                
                Select Case True
                
                    Case aXORBits(xIn + 3, y) = 0
                        
                        aXORBits(xIn + 2, y) = 0
                        aXORBits(xIn + 1, y) = 0
                        aXORBits(xIn + 0, y) = 0
                        
                        xIn = x \ 8
                        aANDBits(xIn, y) = aANDBits(xIn, y) Or m_Pow2(7 - (x Mod 8))
                      
                    Case aXORBits(xIn + 3, y) < 128 And _
                         aXORBits(xIn + 2, y) = 0 And _
                         aXORBits(xIn + 1, y) = 0 And _
                         aXORBits(xIn + 0, y) = 0
                         
                        xIn = x \ 8
                        aANDBits(xIn, y) = aANDBits(xIn, y) Or m_Pow2(7 - (x Mod 8))
                        
                    Case Else
                    
                        xIn = x \ 8
                        aANDBits(xIn, y) = aANDBits(xIn, y) And Not m_Pow2(7 - (x Mod 8))
                End Select
            Next x
        Next y
        
        Call pvUnmapDIB(aXORBits())
        Call pvUnmapDIB(aANDBits())
    End If
End Sub

'//

Friend Sub GetPixelInfo(ByVal x As Long, ByVal y As Long, oXORDIB As cDIB, oANDDIB As cDIB, BPP As dibBPPCts, R As Byte, G As Byte, B As Byte, A As Byte, Index As Byte, IsScreen As Boolean)

  Dim aPixel(3) As Byte
  Dim aPal()    As Byte
  Dim lpXOR     As Long
  Dim lpAND     As Long
  Dim lWXOR     As Long
  Dim lWAND     As Long
  Dim lyRev     As Long
  Dim aMsk      As Byte

    With oXORDIB

        If (x >= 0 And y >= 0 And x < .Width And y < .Height) Then

            lpXOR = .lpBits
            lWXOR = .BytesPerScanline
            lyRev = .Height - y - 1
            lpAND = oANDDIB.lpBits
            lWAND = oANDDIB.BytesPerScanline
            
            '-- Color depth
            BPP = .BPP
            
            '-- Index [?] / RGB color
            Select Case .BPP

                Case [01_bpp]

                    '-- Get palette index
                    Call CopyMemory(aPixel(0), ByVal (lpXOR + lyRev * lWXOR + x \ 8), 1)
                    aMsk = m_Pow2(7 - x Mod 8)
                    Index = (aPixel(0) And aMsk) \ aMsk
                    '-- Get palette color
                    ReDim aPal(7)
                    Call .GetPalette(aPal())
                    R = aPal(4 * Index + 2)
                    G = aPal(4 * Index + 1)
                    B = aPal(4 * Index + 0)

                Case [04_bpp]

                    '-- Get palette index
                    Call CopyMemory(aPixel(0), ByVal (lpXOR + lyRev * lWXOR + x \ 2), 1)
                    If ((x And 1) = 1) Then
                        Index = (aPixel(0) And &HF&)
                      Else
                        Index = (aPixel(0) And &HF0&) \ 16
                    End If
                    '-- Get palette color
                    ReDim aPal(63)
                    Call .GetPalette(aPal())
                    R = aPal(4 * Index + 2)
                    G = aPal(4 * Index + 1)
                    B = aPal(4 * Index + 0)

                Case [08_bpp]

                    '-- Get palette index
                    Call CopyMemory(aPixel(0), ByVal (lpXOR + lyRev * lWXOR + 1 * x), 1)
                    Index = aPixel(0)
                    '-- Get palette color
                    ReDim aPal(1023)
                    Call .GetPalette(aPal())
                    R = aPal(4 * Index + 2)
                    G = aPal(4 * Index + 1)
                    B = aPal(4 * Index + 0)

                Case [24_bpp]

                    '-- Get color
                    Call CopyMemory(aPixel(0), ByVal (lpXOR + lyRev * lWXOR + 3 * x), 3)
                    R = aPixel(2)
                    G = aPixel(1)
                    B = aPixel(0)

                Case [32_bpp]

                    '-- Get color
                    Call CopyMemory(aPixel(0), ByVal (lpXOR + lyRev * lWXOR + 4 * x), 4)
                    A = aPixel(3)
                    R = aPixel(2)
                    G = aPixel(1)
                    B = aPixel(0)
            End Select

            '-- Is screen color [?]
            Call CopyMemory(aPixel(0), ByVal (lpAND + lyRev * lWAND + x \ 8), 1)
            aMsk = m_Pow2(7 - x Mod 8)
            IsScreen = (((aPixel(0) And aMsk) \ aMsk) = 1 And .BPP <> [32_bpp])
        End If
    End With
End Sub

Public Function IsPaletteBlackEntryAvailable(oXORDIB As cDIB, oANDDIB As cDIB, ByVal Index As Byte) As Boolean
    
  Dim bIsLast    As Boolean
  Dim bIsMask    As Boolean
  
  Dim uSAXOR     As SAFEARRAY2D
  Dim uSAAND     As SAFEARRAY2D
  Dim aBitsXOR() As Byte
  Dim aBitsAND() As Byte
  Dim aPal(1023) As Byte
  Dim lIdx       As Long
  Dim lMsk       As Long
  
  Dim x As Long, W As Long
  Dim y As Long, H As Long
    
    Select Case oXORDIB.BPP
    
        Case [01_bpp], [04_bpp]
            
            IsPaletteBlackEntryAvailable = False
            
        Case [08_bpp]
        
            '-- Last black entry [?]
            
            Call oXORDIB.GetPalette(aPal())
            
            bIsLast = True
            For lIdx = 0 To 255
                If (lIdx <> Index) Then
                    If (aPal(4 * lIdx + 0) = 0 And aPal(4 * lIdx + 1) = 0 And aPal(4 * lIdx + 2) = 0) Then
                        bIsLast = False
                        lIdx = 255
                    End If
                End If
            Next lIdx
            
            '-- Used for masking [?]
            
            Call pvMapDIB(oXORDIB, aBitsXOR(), uSAXOR)
            Call pvMapDIB(oANDDIB, aBitsAND(), uSAAND)
            
            W = oXORDIB.Width - 1
            H = oXORDIB.Height - 1
            
            bIsMask = False
            For y = 0 To H
                For x = 0 To W
                    
                    If (aBitsXOR(x, y) = Index) Then
                    
                        lMsk = m_Pow2(7 - x Mod 8)
                        If ((aBitsAND(x \ 8, y) And lMsk) \ lMsk = 1) Then
                            bIsMask = True
                            x = W
                            y = H
                        End If
                    End If
                Next x
            Next y
            
            Call pvUnmapDIB(aBitsXOR())
            Call pvUnmapDIB(aBitsAND())
            
            IsPaletteBlackEntryAvailable = Not (bIsLast Or bIsMask)
            
        Case Else
        
            IsPaletteBlackEntryAvailable = True
    End Select
End Function

'========================================================================================
' Private
'========================================================================================

Private Function pvColorRef(oDIB As cDIB, ByVal Layer As icoLayerCts, ByVal ColorSample As icoColorSample, ByVal A As Byte) As Long
        
  Dim IsScreen As Boolean
  Dim ColorLng As Long
  Dim ColorIdx As Byte
  
    Select Case True
        
        Case (ColorSample = [icoSampleA] And Not m_SwapColors) Or (ColorSample = [icoSampleB] And m_SwapColors)
             
             IsScreen = m_IsAScreen
             ColorLng = m_ColorLngA
             ColorIdx = m_ColorIdxA
        
        Case (ColorSample = [icoSampleB] And Not m_SwapColors) Or (ColorSample = [icoSampleA] And m_SwapColors)
    
             IsScreen = m_IsBScreen
             ColorLng = m_ColorLngB
             ColorIdx = m_ColorIdxB
    End Select
    
    If (IsScreen) Then
        Select Case Layer
            Case [icoXORSolid]
                Select Case oDIB.BPP
                    Case [01_bpp], [04_bpp], [08_bpp], [24_bpp]
                        pvColorRef = &H0
                    Case [32_bpp]
                        pvColorRef = m_ColorScreen
                End Select
            Case [icoXORAlpha]
                pvColorRef = pvARGB(m_ColorScreen, A)
            Case [icoAND]
                pvColorRef = pvDIBINDEX(1)
        End Select
      
      Else
        Select Case Layer
            Case [icoXORSolid]
                Select Case oDIB.BPP
                    Case [01_bpp], [04_bpp], [08_bpp]
                        pvColorRef = pvDIBINDEX(ColorIdx)
                    Case [24_bpp], [32_bpp]
                        pvColorRef = ColorLng
                End Select
            Case [icoXORAlpha]
                pvColorRef = pvARGB(ColorLng, A)
            Case [icoAND]
                pvColorRef = pvDIBINDEX(0)
        End Select
    End If
End Function

Private Function pvDIBINDEX(ByVal Index As Byte) As Long
    
    pvDIBINDEX = pvMakeLong(Index, &H10FF)
End Function

Private Function pvMakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    
    pvMakeLong = ((HiWord * &H10000) + LoWord)
End Function

Private Function pvARGB(ByVal Color As Long, ByVal Alpha As Byte) As Long
  
  Dim uARGB As ARGBQUAD
  Dim aSwap As Byte

   Call CopyMemory(uARGB, Color, 4)
   With uARGB
        .A = Alpha: aSwap = .R: .R = .B: .B = aSwap
   End With
   Call CopyMemory(pvARGB, uARGB, 4)
End Function

Private Function pvRGBToHSL(ByVal R As Long, ByVal G As Long, ByVal B As Long) As HSL
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
    
    With pvRGBToHSL
        
        .L = Lum(lMax)
    
        lDifference = lMax - lMin
        If (lDifference) Then
            
            .S = (lDifference) * 100 / lMax ' do a 65K 2D lookup table here for more speed if needed
            
            q = QTab(lDifference)
            Select Case lMax
                Case R
                    If (B > G) Then
                        .H = q * (G - B) + 360
                      Else
                        .H = q * (G - B)
                    End If
                Case G
                    .H = q * (B - R) + 120
                Case B
                    .H = q * (R - G) + 240
            End Select
        End If
    End With
End Function

Private Sub pvCorrectShapeCoords(x1 As Single, y1 As Single, x2 As Single, y2 As Single)
    
    If (x2 >= x1) Then x2 = x2 + 1 Else If (x2 < x1) Then x1 = x1 + 1
    If (y2 >= y1) Then y2 = y2 + 1 Else If (y2 < y1) Then y1 = y1 + 1
End Sub

Private Sub pvCorrectShapeCoordsGDIp(x1 As Single, y1 As Single, x2 As Single, y2 As Single)

  Dim lSwap As Single
  
    If (x2 < x1) Then lSwap = x1: x1 = x2: x2 = lSwap
    If (y2 < y1) Then lSwap = y1: y1 = y2: y2 = lSwap
End Sub

Private Function pvOLEFontToLogFont(oFont As StdFont, ByVal hDC As Long) As LOGFONT

    With pvOLEFontToLogFont
        
        Call CopyMemory(.lfFaceName(1), ByVal oFont.Name, Len(oFont.Name) + 1)
        .lfCharSet = oFont.Charset
        .lfItalic = -oFont.Italic
        .lfUnderline = -oFont.Underline
        .lfStrikeOut = -oFont.Strikethrough
        .lfWeight = oFont.Weight
        .lfHeight = -(oFont.Size * GetDeviceCaps(hDC, LOGPIXELSY) / 72)
        .lfQuality = NONANTIALIASED_QUALITY
    End With
End Function

'//

Private Sub pvStretchARGB(oSrcDIB32 As cDIB, oDstDIB32 As cDIB)

  Dim hBitmapSrc As Long
  Dim hBitmapDst As Long, hGraphicsDst As Long
  
    Call ARGBBitmapFromGDIDIB32(oSrcDIB32, hBitmapSrc)
    Call ARGBBitmapFromGDIDIB32(oDstDIB32, hBitmapDst)
    Call GdipGetImageGraphicsContext(hBitmapDst, hGraphicsDst)
    Call GdipSetInterpolationMode(hGraphicsDst, [InterpolationModeHighQualityBilinear])
    Call GdipSetPixelOffsetMode(hGraphicsDst, [PixelOffsetModeHighQuality])
    
    Call GdipDrawImageRectRectI(hGraphicsDst, hBitmapSrc, 0, 0, oDstDIB32.Width, oDstDIB32.Height, 0, 0, oSrcDIB32.Width, oSrcDIB32.Height, [UnitPixel], 0, 0, 0)
    
    Call ARGBBitmapToGDIDIB32(oDstDIB32, hBitmapDst)
    
    Call GdipDeleteGraphics(hGraphicsDst)
    Call GdipDisposeImage(hBitmapDst)
    Call GdipDisposeImage(hBitmapSrc)
End Sub

'//

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
