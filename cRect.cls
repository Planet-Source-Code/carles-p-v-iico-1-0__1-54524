VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'================================================
' Class:         cRect.cls
' Author:        Carles P.V.
' Dependencies:  -
' Last revision: 2003.11.02
'================================================

Option Explicit

'-- API:

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT2, lpSrc1Rect As RECT2, lpSrc2Rect As RECT2) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'-- Private Variables:

Private m_RectMain    As RECT2  ' Main rectangle region
Private m_Rect        As RECT2  ' Frame
Private m_RectScaled  As RECT2  ' Frame scaled
Private m_ScaleFactor As Single ' Scale factor

Private m_hBrush(7)   As Long   ' Contour pattern brushes
Private m_nBrush      As Long   ' Current brush



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Initialize()

    '-- Initialize scale factor
    m_ScaleFactor = 1
End Sub

Private Sub Class_Terminate()
    
    '-- Destroy brushes
    Call Me.DestroyPatternBrushes
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub Init(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

    Call SetRect(m_RectMain, x1, y1, x2, y2)
    Call SetRectEmpty(m_Rect)
    Call SetRectEmpty(m_RectScaled)
End Sub

Public Sub Clear()

    Call SetRectEmpty(m_Rect)
    Call SetRectEmpty(m_RectScaled)
End Sub

Public Sub SetCoords(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, Optional ByVal CropMain As Boolean = 0)
    
    Call SetRect(m_Rect, x1, y1, x2, y2)
    Call pvNormalizeRect(m_Rect)
    
    If (CropMain) Then
        Call Me.Crop
      Else
        Call pvScaleRect(m_ScaleFactor)
    End If
End Sub

Public Function GetCoords(x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Boolean

    If (IsRectEmpty(m_Rect) = 0) Then
        With m_Rect
            x1 = .x1
            y1 = .y1
            x2 = .x2
            y2 = .y2
        End With
        GetCoords = True
    End If
End Function

Public Function IsEmpty() As Boolean

    IsEmpty = (IsRectEmpty(m_Rect) <> 0)
End Function

Public Function IsPointIn(ByVal x As Long, ByVal y As Long) As Boolean

    IsPointIn = (PtInRect(m_Rect, x, y) <> 0)
End Function

Public Sub Crop()

    Call IntersectRect(m_Rect, m_Rect, m_RectMain)
    Call pvScaleRect(m_ScaleFactor)
End Sub

Public Sub Offset(ByVal dx As Long, ByVal dy As Long)

    Call OffsetRect(m_Rect, dx, dy)
    Call pvScaleRect(m_ScaleFactor)
End Sub

Public Sub Draw(ByVal hDC As Long, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0)
    
  Dim lOldBkColor As Long
  Dim lOldFrColor As Long
    
    If (IsRectEmpty(m_Rect) = 0 And m_hBrush(m_nBrush) <> 0) Then
    
        '-- Set back = WHITE, fore = BLACK
        lOldBkColor = SetBkColor(hDC, &HFFFFFF)
        lOldFrColor = SetTextColor(hDC, &H0)
        
        '-- Move/Paint/Restore region
        Call OffsetRect(m_RectScaled, xOffset, yOffset)
        Call FrameRect(hDC, m_RectScaled, m_hBrush(m_nBrush))
        Call OffsetRect(m_RectScaled, -xOffset, -yOffset)
        
        '-- Restore back and fore color
        Call SetBkColor(hDC, lOldBkColor)
        Call SetTextColor(hDC, lOldFrColor)
    End If
End Sub

'//

Public Sub CloneTo(oRect As cRect)
    
    Set oRect = New cRect
    Call oRect.Init(m_RectMain.x1, m_RectMain.y1, m_RectMain.x2, m_RectMain.y2)
    Call oRect.SetCoords(m_Rect.x1, m_Rect.y1, m_Rect.x2, m_Rect.y2)
    oRect.ScaleFactor = Me.ScaleFactor
End Sub

'//

Public Sub CreatePatternBrushes()

  Dim lIdx           As Long
  Dim lIdy           As Long
  Dim hBitmap        As Long
  Dim aBytes1(1 To 8) As Integer
  Dim aBytes2(1 To 8) As Integer
    
    '-- Brush main pattern (8x8 monochrome)
    aBytes1(1) = &HF0 '&HCC
    aBytes1(2) = &H78 '&H66
    aBytes1(3) = &H3C '&H33
    aBytes1(4) = &H1E '&H99
    aBytes1(5) = &HF  '&HCC
    aBytes1(6) = &H87 '&H66
    aBytes1(7) = &HC3 '&H33
    aBytes1(8) = &HE1 '&H99
    
    '-- Create main brush
    hBitmap = CreateBitmap(8, 8, 1, 1, aBytes1(1))
    m_hBrush(0) = CreatePatternBrush(hBitmap)
    Call DeleteObject(hBitmap)

    '-- 'Scroll' pattern: Create next brushes
    For lIdx = 7 To 1 Step -1
        For lIdy = 1 To 8
            If (lIdx + lIdy < 9) Then
                aBytes2(lIdy) = aBytes1(lIdx + lIdy)
              Else
                aBytes2(lIdy) = aBytes1(lIdx + lIdy - 8)
            End If
        Next lIdy
        hBitmap = CreateBitmap(8, 8, 1, 1, aBytes2(1))
        m_hBrush(lIdx) = CreatePatternBrush(hBitmap)
        Call DeleteObject(hBitmap)
    Next lIdx
End Sub

Public Sub DestroyPatternBrushes()
    
  Dim lIdx As Long
  
    '-- Destroy pattern brushes
    For lIdx = 0 To 7
        If (m_hBrush(lIdx) <> 0) Then
            Call DeleteObject(m_hBrush(lIdx))
        End If
    Next lIdx
End Sub

Public Sub RotatePatternBrushes()

    m_nBrush = m_nBrush + 1
    If (m_nBrush = 8) Then
        m_nBrush = 0
    End If
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Let ScaleFactor(ByVal Factor As Single)
    m_ScaleFactor = IIf(Factor < 0.001, 0.001, Factor)
    Call pvScaleRect(m_ScaleFactor)
End Property
Public Property Get ScaleFactor() As Single
    ScaleFactor = m_ScaleFactor
End Property

Public Property Get x1() As Long
    x1 = m_Rect.x1
End Property
Public Property Get y1() As Long
    y1 = m_Rect.y1
End Property
Public Property Get x2() As Long
    x2 = m_Rect.x2
End Property
Public Property Get y2() As Long
    y2 = m_Rect.y2
End Property

'========================================================================================
' Private
'========================================================================================

Private Sub pvNormalizeRect(pRect As RECT2)

  Dim lTmp As Long
        
    With pRect
        If (.y2 < .y1) Then
            lTmp = .y2
            .y2 = .y1
            .y1 = lTmp
        End If
        If (.x2 < .x1) Then
            lTmp = .x2
            .x2 = .x1
            .x1 = lTmp
        End If
    End With
End Sub

Private Sub pvScaleRect(ByVal Factor As Single)
    
    '-- Scale frame
    With m_RectScaled
        .x1 = m_Rect.x1 * Factor
        .y1 = m_Rect.y1 * Factor
        .x2 = m_Rect.x2 * Factor + 1
        .y2 = m_Rect.y2 * Factor + 1
    End With
End Sub
