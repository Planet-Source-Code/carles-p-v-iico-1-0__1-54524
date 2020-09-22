VERSION 5.00
Begin VB.UserControl ucIconCanvas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2655
   ClipControls    =   0   'False
   ForeColor       =   &H00808080&
   LockControls    =   -1  'True
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   177
   Begin VB.Timer tmrSelection 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   15
      Top             =   0
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H00808080&
      Height          =   2400
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2400
   End
End
Attribute VB_Name = "ucIconCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucIconCanvas.ctl
' Author:        Carles P.V.
' Dependencies:  cIcon.cls
'                cIconCanvasEx.cls,
'                cDIB.cls
'                cRect.cls
' Last revision: 2004.06.14
'================================================

Option Explicit

'-- API:

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

Private Const PS_SOLID        As Long = 0
Private Const BS_NULL         As Long = 1
Private Const R2_COPYPEN      As Long = 13
Private Const R2_NOT          As Long = 6
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BF_RECT         As Long = &HF

Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function WidenPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function FillPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT2, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)

'//

'-- Public Enums.:
Public Enum icToolKeyCts
    [icSelectionFrame] = 1
    [icPencil] = 2
    [icStraightLine] = 3
    [icBrush] = 4
    [icShape] = 5
    [icFloodFill] = 6
    [icText] = 7
    [icColorSelector] = 8
    [icHotSpot] = 9
End Enum

'-- Default Property Values:
Private Const m_def_ColorScreen As Long = vbWindowBackground
Private Const m_def_ShowGrid    As Boolean = False

'-- Property Variables:
Private m_ShowGrid      As Boolean
Private m_Tool          As icToolKeyCts

'-- Private Constants:
Private Const EDGE3D    As Long = 5

'-- Private Variables:

Private m_lpoIcon       As cIcon   ' Current cIcon object
Private m_ImageIdx      As Integer ' Current cIcon image index
Private m_oCanvasBuffer As cDIB
Private m_oXORBuffer    As cDIB
Private m_oANDBuffer    As cDIB
Private m_oXORClip      As cDIB
Private m_oANDClip      As cDIB
Private m_Copying       As Boolean

Private m_oSrcRect      As cRect
Private m_oDstRect      As cRect
Private m_InRect        As Boolean
Private m_xCur          As Single
Private m_yCur          As Single
Private m_xLst          As Single
Private m_yLst          As Single
Private m_xDwn          As Single
Private m_yDwn          As Single
Private m_InCanvas      As Boolean

Private m_PointsF()     As POINTF
Private m_PointsL()     As POINTL
Private m_hDotBrush     As Long
Private m_ScaleFactor   As Long

'-- Public objects
Public CanvasEx         As cIconCanvasEx

'-- Event Declarations:
Public Event MouseDown(Button As Integer, Shift As Integer, ByVal x As Long, ByVal y As Long, ByVal InSelection As Boolean)
Public Event MouseMove(Button As Integer, Shift As Integer, ByVal x As Long, ByVal y As Long, ByVal InSelection As Boolean)
Public Event MouseUp(Button As Integer, Shift As Integer, ByVal x As Long, ByVal y As Long, ByVal InSelection As Boolean)
Public Event MouseOut()
Public Event IconChange()
Public Event CanvasChange()
Public Event SelectionChange()



'========================================================================================
' UserControl intialization/termination
'========================================================================================

Private Sub UserControl_Initialize()
  
  Dim hBitmap        As Long
  Dim aBytes(1 To 8) As Integer
  Dim lIdx           As Long
    
    '-- Initial scale factor
    m_ScaleFactor = 1
    
    '-- Create dot pattern brush (grid)
    For lIdx = 1 To 8 Step 2
        aBytes(lIdx + 0) = &H55
        aBytes(lIdx + 1) = &HAA
    Next lIdx
    hBitmap = CreateBitmap(8, 8, 1, 1, aBytes(1)): Erase aBytes()
    m_hDotBrush = CreatePatternBrush(hBitmap)
    Call DeleteObject(hBitmap)
    
    '-- Initialize points()
    ReDim m_PointsF(0)
    
    '-- Initialize objects
    '-  Private
    Set m_oCanvasBuffer = New cDIB
    Set m_oXORBuffer = New cDIB
    Set m_oANDBuffer = New cDIB
    Set m_oXORClip = New cDIB
    Set m_oANDClip = New cDIB
    Set m_oSrcRect = New cRect
    Set m_oDstRect = New cRect: Call m_oDstRect.CreatePatternBrushes
    '-  Public
    Set Me.CanvasEx = New cIconCanvasEx
End Sub

Private Sub UserControl_Terminate()
    
    '-- Stop timer [?]
    tmrSelection.Enabled = False
    
    '-- Destroy references-objects
    Set m_lpoIcon = Nothing
    Set m_oXORBuffer = Nothing
    Set m_oANDBuffer = Nothing
    Set m_oCanvasBuffer = Nothing
    Set m_oSrcRect = Nothing
    Set m_oDstRect = Nothing
    Set Me.CanvasEx = Nothing
End Sub

Private Sub UserControl_Show()
    
    '-- Start timer [?]
    tmrSelection.Enabled = Ambient.UserMode
End Sub

'//

Private Sub UserControl_Resize()
               
    On Error Resume Next
    
    Dim W As Long
    Dim H As Long
        
    '-- Get best fit
    With m_oCanvasBuffer
        If ((ScaleWidth + EDGE3D) / .Width > (ScaleHeight + EDGE3D) / .Height) Then
            m_ScaleFactor = (ScaleHeight - EDGE3D) \ .Height
          Else
            m_ScaleFactor = (ScaleWidth - EDGE3D) \ .Width
        End If
        W = .Width * m_ScaleFactor + EDGE3D
        H = .Height * m_ScaleFactor + EDGE3D
    End With
    
    '-- Set selection frame scale factor
    m_oSrcRect.ScaleFactor = m_ScaleFactor
    m_oDstRect.ScaleFactor = m_ScaleFactor
    
    '-- Resize and refresh
    Call picCanvas.Move((ScaleWidth - W) \ 2, (ScaleHeight - H) \ 2, W, H)
    Call Me.Refresh
    
    On Error GoTo 0
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub SetIconSource(oIcon As cIcon)

    '-- Point to source object
    Set m_lpoIcon = oIcon
End Sub

Public Sub SetIconImageIndex(ByVal nImageIdx As Integer)
  
    '-- Store current image index
    m_ImageIdx = nImageIdx
End Sub

'//

Public Sub Initialize()
    
    If (Not m_lpoIcon Is Nothing) Then
    
        '-- Initialize canvas buffer
        Call m_oCanvasBuffer.Create(m_lpoIcon.Width(m_ImageIdx), m_lpoIcon.Height(m_ImageIdx), [32_bpp])
        
        '-- Initialize icon back buffers / selection rects.
        Call pvInitializeIconBackBuffers
        Call pvInitializeSelectionRects
        
        '-- Re-fit
        Call UserControl_Resize
    End If
End Sub

Public Sub Refresh(Optional ByVal DrawToolPointer As Boolean = False)
  
    '-- Refresh canvas
    Call pvRefreshCanvas(DrawToolPointer:=DrawToolPointer)
End Sub

Public Sub PaintCanvas(ByVal hDC As Long, ByVal x As Long, ByVal y As Long)
    
    '-- Paint canvas onto a given hDC
    Call m_oCanvasBuffer.Stretch(hDC, x, y, m_oCanvasBuffer.Width, m_oCanvasBuffer.Height)
End Sub

'//

Public Sub FlipHorizontaly()
    
    With m_lpoIcon
        
        If (m_oDstRect.IsEmpty) Then
            
            '-- Flip icon
            Call Me.CanvasEx.FlipHorizontaly(.oXORDIB(m_ImageIdx))
            If (.BPP(m_ImageIdx) <> [ARGB_Color]) Then
                Call Me.CanvasEx.FlipHorizontaly(.oANDDIB(m_ImageIdx))
            End If
            
            '-- Update buffers / refresh
            Call pvInitializeIconBackBuffers
            Call pvRefreshCanvas(DrawToolPointer:=False)
            
            '-- Notify icon change
            RaiseEvent IconChange
          
          Else
            
            '-- Flip buffer
            Call Me.CanvasEx.FlipHorizontaly(m_oXORBuffer)
            If (.BPP(m_ImageIdx) <> [ARGB_Color]) Then
                Call Me.CanvasEx.FlipHorizontaly(m_oANDBuffer)
            End If
            
            '-- Refresh
            Call pvRefreshCanvas(DrawToolPointer:=False)
        End If
    End With
End Sub

Public Sub FlipVerticaly()

    With m_lpoIcon
        
        If (m_oDstRect.IsEmpty) Then
            
            '-- Flip icon
            Call Me.CanvasEx.FlipVerticaly(.oXORDIB(m_ImageIdx))
            If (.BPP(m_ImageIdx) <> [ARGB_Color]) Then
                Call Me.CanvasEx.FlipVerticaly(.oANDDIB(m_ImageIdx))
            End If
            
            '-- Update buffers / refresh
            Call pvInitializeIconBackBuffers
            Call pvRefreshCanvas(DrawToolPointer:=False)
            
            '-- Notify icon change
            RaiseEvent IconChange
          
          Else
            
            '-- Flip buffer
            Call Me.CanvasEx.FlipVerticaly(m_oXORBuffer)
            If (.BPP(m_ImageIdx) <> [ARGB_Color]) Then
                Call Me.CanvasEx.FlipVerticaly(m_oANDBuffer)
            End If
            
            '-- Refresh
            Call pvRefreshCanvas(DrawToolPointer:=False)
        End If
    End With
End Sub

Public Sub RotateLeft()

    With m_lpoIcon
        
        If (m_oDstRect.IsEmpty) Then
            
            '-- Rotate icon
            Call Me.CanvasEx.RotateLeft(.oXORDIB(m_ImageIdx))
            If (.BPP(m_ImageIdx) <> [ARGB_Color]) Then
                Call Me.CanvasEx.RotateLeft(.oANDDIB(m_ImageIdx))
            End If
            
            '-- Update buffers
            Call pvInitializeIconBackBuffers
            
            '-- Initialize canvas buffer / Re-fit canvas (and refresh)
            Call m_oCanvasBuffer.Create(m_lpoIcon.Width(m_ImageIdx), m_lpoIcon.Height(m_ImageIdx), [32_bpp])
            Call UserControl_Resize
            
            '-- Notify icon change
            RaiseEvent IconChange
          
          Else
            
            '-- Rotate buffer
            Call Me.CanvasEx.RotateLeft(m_oXORBuffer)
            If (.BPP(m_ImageIdx) <> [ARGB_Color]) Then
                Call Me.CanvasEx.RotateLeft(m_oANDBuffer)
            End If
            '-- And selection rect
            With m_oDstRect
                Call .SetCoords(.x1, .y1, .x1 + (.y2 - .y1), .y1 + (.x2 - .x1))
                '-- Check if out of canvas...
                If (.x2 < 1) Then Call .Offset(-.x2 + 1, 0)
                If (.y2 < 1) Then Call .Offset(0, -.y2 + 1)
            End With
            
            '-- Refresh
            Call pvRefreshCanvas(DrawToolPointer:=False)
        End If
    End With
End Sub

Public Sub RotateRight()

    With m_lpoIcon
        
        If (m_oDstRect.IsEmpty) Then
            
            '-- Rotate icon
            Call Me.CanvasEx.RotateRight(.oXORDIB(m_ImageIdx))
            If (.BPP(m_ImageIdx) <> [ARGB_Color]) Then
                Call Me.CanvasEx.RotateRight(.oANDDIB(m_ImageIdx))
            End If
            
            '-- Update buffers
            Call pvInitializeIconBackBuffers
            
            '-- Initialize canvas buffer / Re-fit canvas (and refresh)
            Call m_oCanvasBuffer.Create(m_lpoIcon.Width(m_ImageIdx), m_lpoIcon.Height(m_ImageIdx), [32_bpp])
            Call UserControl_Resize
            
            '-- Notify icon change
            RaiseEvent IconChange
          
          Else
            
            '-- Rotate buffer
            Call Me.CanvasEx.RotateRight(m_oXORBuffer)
            If (.BPP(m_ImageIdx) <> [ARGB_Color]) Then
                Call Me.CanvasEx.RotateRight(m_oANDBuffer)
            End If
            '-- And selection rect
            With m_oDstRect
                Call .SetCoords(.x1, .y1, .x1 + (.y2 - .y1), .y1 + (.x2 - .x1))
                '-- Check if out of canvas...
                If (.x2 < 1) Then Call .Offset(-.x2 + 1, 0)
                If (.y2 < 1) Then Call .Offset(0, -.y2 + 1)
            End With
            
            '-- Refresh
            Call pvRefreshCanvas(DrawToolPointer:=False)
        End If
    End With
End Sub

'//

Public Sub SelectAll()
    
    If (Not m_lpoIcon Is Nothing) Then
        
        With m_lpoIcon
            
            '-- Change current tool
            m_Tool = [icSelectionFrame]
            picCanvas.MouseIcon = LoadResPicture("CURSOR_MOVE", vbResCursor)
            
            '-- Set rects.
            Call m_oDstRect.Init(0, 0, .Width(m_ImageIdx), .Height(m_ImageIdx))
            Call m_oDstRect.SetCoords(0, 0, .Width(m_ImageIdx), .Height(m_ImageIdx))
            Call m_oDstRect.CloneTo(m_oSrcRect)
            
            '-- Update buffers / refresh
            Call pvInitializeIconBackBuffers
            Call pvRefreshCanvas(DrawToolPointer:=False)
            
            RaiseEvent SelectionChange
        End With
    End If
End Sub

Public Sub ClearSelection()

    If (Not m_lpoIcon Is Nothing) Then
        
        If (m_Tool = [icSelectionFrame] And Not m_oDstRect.IsEmpty) Then
        
            '-- Clear rectangle and initialize canvas
            With m_lpoIcon
                Call Me.CanvasEx.ClearIconRect(m_oDstRect, m_lpoIcon.oXORDIB(m_ImageIdx), m_lpoIcon.oANDDIB(m_ImageIdx))
                Call Me.Initialize
            End With
            picCanvas.MouseIcon = LoadResPicture("CURSOR_FRAMESELECTION", vbResCursor)
        
            '-- Refresh
            Call pvRefreshCanvas(DrawToolPointer:=False)
        
            '-- Notify icon/selection change
            RaiseEvent IconChange
            RaiseEvent SelectionChange
        End If
    End If
End Sub

Public Sub Cut()

    If (Not m_lpoIcon Is Nothing) Then
            
        If (m_Tool = [icSelectionFrame] And Not m_oDstRect.IsEmpty) Then
        
            '-- Copy to Clipboard, clear rectangle and initialize canvas
            With m_lpoIcon
                Call Me.Copy
                Call Me.CanvasEx.ClearIconRect(m_oSrcRect, .oXORDIB(m_ImageIdx), .oANDDIB(m_ImageIdx))
                Call Me.Initialize
            End With
            picCanvas.MouseIcon = LoadResPicture("CURSOR_FRAMESELECTION", vbResCursor)
                
            '-- Refresh
            Call pvRefreshCanvas(DrawToolPointer:=False)
            
            '-- Notify icon/selection change
            RaiseEvent IconChange
            RaiseEvent SelectionChange
        End If
    End If
End Sub

Public Sub Copy()
    
    '-- Make private copy
    Call m_oXORBuffer.CloneTo(m_oXORClip)
    Call m_oANDBuffer.CloneTo(m_oANDClip)
    
    '-- Export XOR buffer to Clipboard
    m_Copying = True
    Call Me.CanvasEx.ExportIconToWindowsClipboard(m_oXORBuffer, m_oANDBuffer)
End Sub

Public Sub DestroyPrivateClipboard()
    
    If (Not m_Copying) Then
        Call m_oXORClip.Destroy
        Call m_oANDClip.Destroy
      Else
        m_Copying = False
    End If
End Sub

Public Sub Paste()
    
  Dim bProcess As Boolean
  
    If (m_oXORClip.hDIB) Then
        Call Me.CanvasEx.ImportIconFromPrivateClipboard(m_oXORBuffer, m_oANDBuffer, m_oXORClip, m_oANDClip, m_oDstRect)
        bProcess = True
      Else
        If (Clipboard.GetFormat(vbCFDIB)) Then
            Call Me.CanvasEx.ImportIconFromWindowsClipboard(m_oXORBuffer, m_oANDBuffer, m_oDstRect)
            bProcess = True
        End If
    End If
    
    If (bProcess) Then
    
        '-- Change current tool
        m_Tool = [icSelectionFrame]
        picCanvas.MouseIcon = LoadResPicture("CURSOR_MOVE", vbResCursor)
        
        '-- Initialize buffers/rects.
        If (m_oSrcRect.x1 = m_oDstRect.x1 And _
            m_oSrcRect.y1 = m_oDstRect.y1 And _
            m_oSrcRect.x2 = m_oDstRect.x2 And _
            m_oSrcRect.y2 = m_oDstRect.y2) Then
            Call m_oSrcRect.Clear
        End If
        
        '-- Refresh
        Call pvRefreshCanvas(DrawToolPointer:=False)
        
        RaiseEvent SelectionChange
    End If
End Sub

Public Sub MergeSelection()
 
  Dim oImageRect As cRect
    
    Select Case m_Tool
    
        Case [icSelectionFrame]
        
            Set oImageRect = New cRect
            Call oImageRect.SetCoords(0, 0, m_oXORBuffer.Width, m_oXORBuffer.Height)
            
            With m_lpoIcon
                Call Me.CanvasEx.ClearIconRect(m_oSrcRect, .oXORDIB(m_ImageIdx), .oANDDIB(m_ImageIdx))
                Call Me.CanvasEx.MergeIcon(m_oDstRect, oImageRect, .oXORDIB(m_ImageIdx), .oANDDIB(m_ImageIdx), m_oXORBuffer, m_oANDBuffer)
            End With
            
            Call m_oSrcRect.Clear
                    
            '-- Refresh
            Call pvRefreshCanvas(DrawToolPointer:=False)
            
            '-- Notify icon change
            RaiseEvent IconChange
        
        Case [icText]
            
            Call pvDrawTextLayer(ProcessColorScreen:=Me.CanvasEx.IsAScreen)
            
            '-- Update buffers / refresh
            Call pvInitializeIconBackBuffers
            Call pvRefreshCanvas(DrawToolPointer:=False)
            
            '-- Notify icon change
            RaiseEvent IconChange
    End Select
End Sub

'//

Public Sub GetMouseInfo(xDwn As Long, yDwn As Long, xCur As Long, yCur As Long)
    
    If (Not m_lpoIcon Is Nothing) Then
        
        xDwn = Int(m_xDwn)
        yDwn = Int(m_yDwn)
        xCur = Int(m_xCur)
        yCur = Int(m_yCur)
    End If
End Sub

Public Sub GetPixelInfo(ByVal x As Long, ByVal y As Long, BPP As dibBPPCts, R As Byte, G As Byte, B As Byte, A As Byte, Index As Byte, IsScreen As Boolean)
    
    If (Not m_lpoIcon Is Nothing) Then
        
        Call Me.CanvasEx.GetPixelInfo _
            (x, y, _
             m_lpoIcon.oXORDIB(m_ImageIdx), _
             m_lpoIcon.oANDDIB(m_ImageIdx), _
             BPP, _
             R, G, B, A, _
             Index, _
             IsScreen _
            )
    End If
End Sub

Public Sub GetSelectionInfo(x As Long, y As Long, W As Long, H As Long)
    
    If (Not m_lpoIcon Is Nothing) Then
        
        With m_oDstRect
            x = .x1
            y = .y1
            W = .x2 - .x1
            H = .y2 - .y1
        End With
    End If
End Sub

'========================================================================================
' UserControl paint processing
'========================================================================================

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim rctControl As RECT2
  Dim rctOffset  As RECT2
  
    m_xDwn = pvDIBx(x)
    m_yDwn = pvDIBx(y)
    m_xLst = m_xCur
    m_yLst = m_yCur
    
    m_InCanvas = (x >= EDGE3D \ 2 And x < picCanvas.Width - EDGE3D \ 2 - 1 And y >= EDGE3D \ 2 And y < picCanvas.Height - EDGE3D \ 2 - 1)
    
    If (m_InCanvas) Then
        
        Me.CanvasEx.SwapColors = (Button = vbRightButton)
        
        Select Case m_Tool
            
            Case [icSelectionFrame]
                
                m_InRect = (m_oDstRect.IsPointIn(Int(m_xDwn), Int(m_yDwn)))
                
                If (Button = vbLeftButton And Not m_InRect) Then
                    Call m_oDstRect.Clear
                    Call m_oSrcRect.Clear
                    RaiseEvent SelectionChange
                End If
                Call GetWindowRect(picCanvas.hWnd, rctOffset)
                Call GetClientRect(picCanvas.hWnd, rctControl)
                Call OffsetRect(rctControl, rctOffset.x1, rctOffset.y1)
                Call InflateRect(rctControl, -(EDGE3D \ 2 + 1), -(EDGE3D \ 2 + 1))
                Call ClipCursor(rctControl)
                
                Call picCanvas_MouseMove(Button, Shift, x, y)
                
            Case [icPencil]
                
                ReDim m_PointsL(0)
                m_PointsL(0).x = Int(m_xDwn)
                m_PointsL(0).y = Int(m_yDwn)
                
                Call picCanvas_MouseMove(Button, Shift, x, y)
            
            Case [icStraightLine]
               
                Call picCanvas_MouseMove(Button, Shift, x, y)
                
            Case [icBrush]
                    
                ReDim m_PointsF(0)
                m_PointsF(0).x = m_xDwn
                m_PointsF(0).y = m_yDwn
                
                Call picCanvas_MouseMove(Button, Shift, x, y)
            
            Case [icShape]
            
                Call picCanvas_MouseMove(Button, Shift, x, y)
            
            Case [icFloodFill]
                
                Call Me.CanvasEx.FloodFill(m_lpoIcon.oXORDIB(m_ImageIdx), m_lpoIcon.oANDDIB(m_ImageIdx), Int(m_xCur), Int(m_yCur))
                Call pvRefreshCanvas(DrawToolPointer:=False)
                
                '-- Update buffers
                Call pvInitializeIconBackBuffers
                
                '-- Notify icon change
                RaiseEvent IconChange
                           
            Case [icText]
            
                m_InRect = (m_oDstRect.IsPointIn(Int(m_xDwn), Int(m_yDwn)))
                
                Call GetWindowRect(picCanvas.hWnd, rctOffset)
                Call GetClientRect(picCanvas.hWnd, rctControl)
                Call OffsetRect(rctControl, rctOffset.x1, rctOffset.y1)
                Call InflateRect(rctControl, -(EDGE3D \ 2 + 1), -(EDGE3D \ 2 + 1))
                Call ClipCursor(rctControl)
                
            Case [icColorSelector]
            
            Case [icHotSpot]
            
        End Select
        
        RaiseEvent MouseDown(Button, Shift, Int(m_xCur), Int(m_yCur), m_oDstRect.IsPointIn(Int(m_xCur), Int(m_yCur)))
    End If
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    m_xCur = pvDIBx(x)
    m_yCur = pvDIBx(y)
    
    Select Case m_Tool
        Case [icSelectionFrame]
        
            If (m_InCanvas And Button = vbLeftButton) Then
                    
                '-- Move selection [?]
                If (m_InRect) Then
                    Call m_oDstRect.Offset(Int(m_xCur) - Int(m_xLst), Int(m_yCur) - Int(m_yLst))
                End If
                '-- Refresh
                Call pvRefreshCanvas(DrawToolPointer:=Not m_InRect)
              
              Else
                '-- Update pointer
                If (m_oDstRect.IsPointIn(Int(m_xCur), Int(m_yCur))) Then
                    picCanvas.MouseIcon = LoadResPicture("CURSOR_MOVE", vbResCursor)
                  Else
                    picCanvas.MouseIcon = LoadResPicture("CURSOR_FRAMESELECTION", vbResCursor)
                End If
            End If
    
        Case [icPencil]
        
            If (m_InCanvas And Button) Then
               
                '-- Add new point
                ReDim Preserve m_PointsL(UBound(m_PointsL()) + 1)
                m_PointsL(UBound(m_PointsL())).x = Int(m_xCur)
                m_PointsL(UBound(m_PointsL())).y = Int(m_yCur)
                
                '-- Restore icon / draw pixels / refresh
                Call pvRestoreIcon
                Call Me.CanvasEx.DrawPixels(m_lpoIcon.oXORDIB(m_ImageIdx), m_PointsL(), -(m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color]))
                If (m_lpoIcon.oXORDIB(m_ImageIdx).BPP <> [32_bpp]) Then
                    Call Me.CanvasEx.DrawPixels(m_lpoIcon.oANDDIB(m_ImageIdx), m_PointsL(), [icoAND])
                End If
            End If
            Call pvRefreshCanvas(DrawToolPointer:=(Button = 0))
        
        Case [icStraightLine]

            If (m_InCanvas And Button) Then
            
                '-- Set line points
                ReDim m_PointsF(1)
                m_PointsF(0).x = m_xDwn
                m_PointsF(0).y = m_yDwn
                m_PointsF(1).x = m_xCur
                m_PointsF(1).y = m_yCur
                
                '-- Restore icon / draw straight line / refresh
                Call pvRestoreIcon
                Call Me.CanvasEx.DrawStraightLine(m_lpoIcon.oXORDIB(m_ImageIdx), m_PointsF(0).x, m_PointsF(0).y, m_PointsF(1).x, m_PointsF(1).y, -(m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color]))
                If (m_lpoIcon.oXORDIB(m_ImageIdx).BPP <> [32_bpp]) Then
                    Call Me.CanvasEx.DrawStraightLine(m_lpoIcon.oANDDIB(m_ImageIdx), m_PointsF(0).x, m_PointsF(0).y, m_PointsF(1).x, m_PointsF(1).y, [icoAND])
                End If
            End If
            Call pvRefreshCanvas(DrawToolPointer:=(Button = 0))
            
        Case [icBrush]
            
            If (m_InCanvas And Button) Then
            
                '-- Add new line point
                ReDim Preserve m_PointsF(UBound(m_PointsF()) + 1)
                m_PointsF(UBound(m_PointsF())).x = m_xCur
                m_PointsF(UBound(m_PointsF())).y = m_yCur
                
                '-- Restore icon / draw brush lines / refresh
                Call pvRestoreIcon
                Call Me.CanvasEx.DrawBrushLines(m_lpoIcon.oXORDIB(m_ImageIdx), m_PointsF(), -(m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color]))
                If (m_lpoIcon.oXORDIB(m_ImageIdx).BPP <> [32_bpp]) Then
                    Call Me.CanvasEx.DrawBrushLines(m_lpoIcon.oANDDIB(m_ImageIdx), m_PointsF(), [icoAND])
                End If
            End If
            Call pvRefreshCanvas(DrawToolPointer:=(Button = 0))
            
        Case [icShape]
        
            If (m_InCanvas And Button) Then
            
                '-- Restore icon / draw shape / refresh
                Call pvRestoreIcon
                Call Me.CanvasEx.DrawShape(m_lpoIcon.oXORDIB(m_ImageIdx), m_xDwn, m_yDwn, m_xCur, m_yCur, -(m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color]))
                If (m_lpoIcon.oXORDIB(m_ImageIdx).BPP <> [32_bpp]) Then
                    Call Me.CanvasEx.DrawShape(m_lpoIcon.oANDDIB(m_ImageIdx), m_xDwn, m_yDwn, m_xCur, m_yCur, [icoAND])
                End If
            End If
            Call pvRefreshCanvas(DrawToolPointer:=(Button = 0))
        
        Case [icFloodFill]
            
        Case [icText]
        
            If (m_InCanvas And Button = vbLeftButton) Then
                    
                '-- Move selection [?]
                If (m_InRect) Then
                    Call m_oDstRect.Offset(Int(m_xCur) - Int(m_xLst), Int(m_yCur) - Int(m_yLst))
                    Call pvRefreshCanvas(DrawToolPointer:=False)
                End If
              
              Else
                '-- Update pointer
                If (m_oDstRect.IsPointIn(Int(m_xCur), Int(m_yCur))) Then
                    picCanvas.MouseIcon = LoadResPicture("CURSOR_MOVE", vbResCursor)
                  Else
                    picCanvas.MouseIcon = LoadResPicture("CURSOR_NODROP", vbResCursor)
                End If
            End If
            
        Case [icText]
            
        Case [icColorSelector]
        
        Case [icHotSpot]
    
    End Select
    
    m_xLst = m_xCur
    m_yLst = m_yCur
    RaiseEvent MouseMove(Button, Shift, Int(m_xCur), Int(m_yCur), m_oDstRect.IsPointIn(Int(m_xCur), Int(m_yCur)))
    
    If (Button = 0) Then
        If (GetCapture <> picCanvas.hWnd And Button = 0) Then
            Call SetCapture(picCanvas.hWnd)
        End If
        If (x < 0 Or y < 0 Or x >= picCanvas.Width Or y >= picCanvas.Height) Then
            m_xLst = -1
            m_yLst = -1
            Call pvRefreshCanvas(DrawToolPointer:=False)
            Call ReleaseCapture
            RaiseEvent MouseOut
        End If
    End If
End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 
  Dim bProcessColorScreen As Boolean
  Dim oAlpha              As cDIB
  
    If (m_InCanvas) Then
        m_InCanvas = False
        
        bProcessColorScreen = (Me.CanvasEx.IsAScreen And Not Me.CanvasEx.SwapColors) Or _
                              (Me.CanvasEx.IsBScreen And Me.CanvasEx.SwapColors)
        
        Select Case m_Tool
            
            Case [icSelectionFrame]
                
                Call ClipCursor(ByVal 0)
                
                If (Button = vbLeftButton And m_oDstRect.IsEmpty) Then
                    
                    '-- Set new selection rect.
                    Call m_oDstRect.SetCoords(Int(m_xDwn) - (m_xCur < m_xDwn), _
                                              Int(m_yDwn) - (m_yCur < m_yDwn), _
                                              Int(m_xCur) - (m_xCur >= m_xDwn), _
                                              Int(m_yCur) - (m_yCur >= m_yDwn))
                    Call m_oDstRect.Crop
                    Call m_oDstRect.CloneTo(m_oSrcRect)
                    
                    If (m_oDstRect.IsEmpty = False) Then
                        
                        '-- Update buffers
                        Call pvInitializeIconBackBuffers
                        Call Me.CanvasEx.CropIcon(m_oDstRect, m_oXORBuffer, m_oANDBuffer)
                        '-- Refresh
                        Call pvRefreshCanvas(DrawToolPointer:=False)
                        
                        picCanvas.MouseIcon = LoadResPicture("CURSOR_MOVE", vbResCursor)
                    End If
                    RaiseEvent SelectionChange
                End If
                
            Case [icPencil]
            
                If (bProcessColorScreen) Then
                    Set oAlpha = New cDIB
                    With m_lpoIcon.oXORDIB(m_ImageIdx)
                        Call oAlpha.Create(.Width, .Height, [32_bpp])
                        Call oAlpha.Reset
                    End With
                    Call Me.CanvasEx.DrawPixels(oAlpha, m_PointsL(), [icoXORAlpha])
                    Call Me.CanvasEx.ProcessColorScreen(m_lpoIcon.oXORDIB(m_ImageIdx), m_oXORBuffer, oAlpha)
                End If
                Call pvRefreshCanvas(DrawToolPointer:=(m_oXORBuffer.BPP = [32_bpp]))
                
                '-- Update buffers
                Call pvInitializeIconBackBuffers
                
                '-- Notify icon change
                RaiseEvent IconChange
                
            Case [icStraightLine]
            
                If (bProcessColorScreen) Then
                    Set oAlpha = New cDIB
                    With m_lpoIcon.oXORDIB(m_ImageIdx)
                        Call oAlpha.Create(.Width, .Height, [32_bpp])
                        Call oAlpha.Reset
                    End With
                    Call Me.CanvasEx.DrawStraightLine(oAlpha, m_PointsF(0).x, m_PointsF(0).y, m_PointsF(1).x, m_PointsF(1).y, [icoXORAlpha])
                    Call Me.CanvasEx.ProcessColorScreen(m_lpoIcon.oXORDIB(m_ImageIdx), m_oXORBuffer, oAlpha)
                End If
                Call pvRefreshCanvas(DrawToolPointer:=(m_oXORBuffer.BPP = [32_bpp]))
                
                '-- Update buffers
                Call pvInitializeIconBackBuffers
                
                '-- Notify icon change
                RaiseEvent IconChange
           
            Case [icBrush]
                
                If (bProcessColorScreen) Then
                    Set oAlpha = New cDIB
                    With m_lpoIcon.oXORDIB(m_ImageIdx)
                        Call oAlpha.Create(.Width, .Height, [32_bpp])
                        Call oAlpha.Reset
                   End With
                    Call Me.CanvasEx.DrawBrushLines(oAlpha, m_PointsF(), [icoXORAlpha])
                    Call Me.CanvasEx.ProcessColorScreen(m_lpoIcon.oXORDIB(m_ImageIdx), m_oXORBuffer, oAlpha)
                End If
                Call pvRefreshCanvas(DrawToolPointer:=(m_oXORBuffer.BPP = [32_bpp]))
                
                 '-- Update buffers
                Call pvInitializeIconBackBuffers
                
                '-- Notify icon change
                RaiseEvent IconChange
           
            Case [icShape]
                
                If (bProcessColorScreen) Then
                    Set oAlpha = New cDIB
                    With m_lpoIcon.oXORDIB(m_ImageIdx)
                        Call oAlpha.Create(.Width, .Height, [32_bpp])
                        Call oAlpha.Reset
                    End With
                    Call Me.CanvasEx.DrawShape(oAlpha, m_xDwn, m_yDwn, m_xCur, m_yCur, [icoXORAlpha])
                    Call Me.CanvasEx.ProcessColorScreen(m_lpoIcon.oXORDIB(m_ImageIdx), m_oXORBuffer, oAlpha)
                End If
                Call pvRefreshCanvas(DrawToolPointer:=(m_oXORBuffer.BPP = [32_bpp]))
                
                '-- Update buffers
                Call pvInitializeIconBackBuffers
                
                '-- Notify icon change
                RaiseEvent IconChange
                
            Case [icText]
            
                Call ClipCursor(ByVal 0)
    
            Case [icFloodFill]
            
            Case [icColorSelector]
            
            Case [icHotSpot]
            
        End Select
        
        RaiseEvent MouseUp(Button, Shift, Int(m_xCur), Int(m_yCur), m_oDstRect.IsPointIn(Int(m_xCur), Int(m_yCur)))
    End If
    
    If (x < 0 Or y < 0 Or x >= picCanvas.Width Or y >= picCanvas.Height) Then
        Call pvRefreshCanvas(DrawToolPointer:=False)
        RaiseEvent MouseOut
    End If
End Sub

Private Sub tmrSelection_Timer()
    
    If (Not m_oDstRect.IsEmpty) Then
        Call m_oDstRect.RotatePatternBrushes
        Call m_oDstRect.Draw(picCanvas.hDC, EDGE3D \ 2, EDGE3D \ 2)
        Call picCanvas.Refresh
    End If
    If (m_Tool = [icHotSpot]) Then
        Call pvRefreshCanvas(DrawToolPointer:=True)
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvRestoreIcon()
    
    If (Not m_lpoIcon Is Nothing) Then
        Call CopyMemory(ByVal m_lpoIcon.oXORDIB(m_ImageIdx).lpBits, ByVal m_oXORBuffer.lpBits, m_oXORBuffer.Size)
        Call CopyMemory(ByVal m_lpoIcon.oANDDIB(m_ImageIdx).lpBits, ByVal m_oANDBuffer.lpBits, m_oANDBuffer.Size)
    End If
End Sub

Private Sub pvInitializeIconBackBuffers()
  
    If (Not m_lpoIcon Is Nothing) Then
        With m_lpoIcon
            Call .oXORDIB(m_ImageIdx).CloneTo(m_oXORBuffer)
            Call .oANDDIB(m_ImageIdx).CloneTo(m_oANDBuffer)
        End With
    End If
End Sub

Private Sub pvInitializeSelectionRects()
    
    If (Not m_lpoIcon Is Nothing) Then
        With m_lpoIcon
            Call m_oSrcRect.Init(0, 0, .Width(m_ImageIdx), .Height(m_ImageIdx))
            Call m_oDstRect.Init(0, 0, .Width(m_ImageIdx), .Height(m_ImageIdx))
            RaiseEvent SelectionChange
        End With
    End If
End Sub

Private Sub pvRefreshCanvas(Optional ByVal DrawToolPointer As Boolean = False)
 
    If (Extender.Visible) Then
 
        If (Not m_lpoIcon Is Nothing) Then
            
            With m_oCanvasBuffer
                
                Call .Cls(Me.CanvasEx.ColorScreen)
                
                Select Case m_Tool
                
                    Case [icSelectionFrame]
                    
                        Call m_lpoIcon.DrawIconStretch(m_ImageIdx, .hDC)
                        Call pvDrawFrameLayer
                    
                    Case [icText]
                        
                        Call pvRestoreIcon
                        Call pvDrawTextLayer(ProcessColorScreen:=False)
                        Call m_lpoIcon.DrawIconStretch(m_ImageIdx, .hDC)
                        Call pvRestoreIcon
                    
                    Case Else
                    
                        Call m_lpoIcon.DrawIconStretch(m_ImageIdx, .hDC)
                End Select
                
                If (DrawToolPointer) Then
                    Call pvDrawToolPointer
                End If
                        
                Call .Stretch(picCanvas.hDC, 2, 2, m_ScaleFactor * .Width, m_ScaleFactor * .Height)
            End With
            
            Call pvDrawDotGrid
            Call m_oDstRect.Draw(picCanvas.hDC, EDGE3D \ 2, EDGE3D \ 2)
            
            Call picCanvas.Refresh
        End If
        
        RaiseEvent CanvasChange
    End If
End Sub

'//

Private Sub pvDrawFrameLayer()

  Dim rctErase As RECT2
  Dim hBrush   As Long
                        
    If (Not m_lpoIcon Is Nothing) Then
    
        If (Not m_oDstRect.IsEmpty) Then
            
            '-- Erase source background
            hBrush = CreateSolidBrush(Me.CanvasEx.ColorScreen)
            Call m_oSrcRect.GetCoords(rctErase.x1, rctErase.y1, rctErase.x2, rctErase.y2)
            Call FillRect(m_oCanvasBuffer.hDC, rctErase, hBrush)
            Call DeleteObject(hBrush)
                
            '-- Draw frame
            With m_oDstRect
                If (m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color]) Then
                    Call m_oXORBuffer.Stretch32(m_oCanvasBuffer.hDC, .x1, .y1, .x2 - .x1, .y2 - .y1, , , , , Me.CanvasEx.Alpha)
                  Else
                    Call m_oANDBuffer.Stretch(m_oCanvasBuffer.hDC, .x1, .y1, .x2 - .x1, .y2 - .y1, , , , , vbSrcAnd)
                    Call m_oXORBuffer.Stretch(m_oCanvasBuffer.hDC, .x1, .y1, .x2 - .x1, .y2 - .y1, , , , , vbSrcPaint)
                End If
            End With
        End If
    End If
End Sub

Private Sub pvDrawTextLayer(Optional ProcessColorScreen As Boolean = False)

  Dim oAlpha As cDIB
  Dim W As Single, H As Single
    
    If (Not m_oDstRect Is Nothing) Then
    
        With m_oDstRect
        
            '-- Calc. text rect.
            Call Me.CanvasEx.MeasureString(m_lpoIcon.oXORDIB(m_ImageIdx), W, H, m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color])
            Call .SetCoords(.x1, .y1, .x1 + W, .y1 + H)
            '-- Check if out of canvas...
            If (.x1 < 0 And .x2 < 1) Then Call .Offset(-.x2 + 1, 0)
            If (.y1 < 0 And .y2 < 1) Then Call .Offset(0, -.y2 + 1)
                
            '-- Draw text
            Me.CanvasEx.SwapColors = False
            Call Me.CanvasEx.DrawString(m_lpoIcon.oXORDIB(m_ImageIdx), .x1, .y1, W, H, -(m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color]))
            If (m_lpoIcon.oXORDIB(m_ImageIdx).BPP <> [32_bpp]) Then
                Call Me.CanvasEx.DrawString(m_lpoIcon.oANDDIB(m_ImageIdx), .x1, .y1, W, H, [icoAND])
            End If
            
            '-- Process color screen [?]
            If (ProcessColorScreen And m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color]) Then
                Set oAlpha = New cDIB
                With m_lpoIcon.oXORDIB(m_ImageIdx)
                    Call oAlpha.Create(m_lpoIcon.oXORDIB(m_ImageIdx).Width, m_lpoIcon.oXORDIB(m_ImageIdx).Height, [32_bpp])
                    Call oAlpha.Reset
                End With
                Call Me.CanvasEx.DrawString(oAlpha, .x1, .y1, W, H, [icoXORAlpha])
                Call Me.CanvasEx.ProcessColorScreen(m_lpoIcon.oXORDIB(m_ImageIdx), m_oXORBuffer, oAlpha)
            End If
        End With
    End If
End Sub

Private Sub pvDrawToolPointer()

  Dim oDIB32        As New cDIB
  Dim ptPointerF(1) As POINTF
  Dim ptPointerL(1) As POINTL
                        
    If (Not m_lpoIcon Is Nothing And GetCapture = picCanvas.hWnd) Then
    
        With m_oCanvasBuffer
        
            ptPointerF(0).x = m_xCur
            ptPointerF(0).y = m_yCur
            ptPointerF(1) = ptPointerF(0)
            ptPointerL(0).x = Int(m_xCur)
            ptPointerL(0).y = Int(m_yCur)
            ptPointerL(1) = ptPointerL(0)
            
            Select Case m_Tool
                
                Case [icSelectionFrame]
                
                    Call SetROP2(.hDC, R2_NOT)
                    Call pvDrawRect(.hDC, Int(m_xDwn), Int(m_yDwn), Int(m_xCur), Int(m_yCur))
                    Call SetROP2(.hDC, R2_COPYPEN)
                    
                Case [icPencil]
                     
                    If (m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color]) Then
                        Call oDIB32.Create(.Width, .Height, [32_bpp])
                        Call oDIB32.Reset
                        Call Me.CanvasEx.DrawPixels(oDIB32, ptPointerL(), [icoXORAlpha])
                        Call oDIB32.Stretch32(.hDC, 0, 0, .Width, .Height)
                      Else
                        Call Me.CanvasEx.DrawPixels(m_oCanvasBuffer, ptPointerL(), [icoXORSolid])
                    End If
                    
                Case [icStraightLine]
                
                    If (m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color]) Then
                        Call oDIB32.Create(.Width, .Height, [32_bpp])
                        Call oDIB32.Reset
                        Call Me.CanvasEx.DrawStraightLine(oDIB32, ptPointerF(0).x, ptPointerF(0).y, ptPointerF(1).x, ptPointerF(1).y, [icoXORAlpha])
                        Call oDIB32.Stretch32(.hDC, 0, 0, .Width, .Height)
                      Else
                        Call Me.CanvasEx.DrawStraightLine(m_oCanvasBuffer, ptPointerF(0).x, ptPointerF(0).y, ptPointerF(1).x, ptPointerF(1).y, [icoXORSolid])
                    End If
                    
                Case [icBrush]
                
                    If (m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color]) Then
                        Call oDIB32.Create(.Width, .Height, [32_bpp])
                        Call oDIB32.Reset
                        Call Me.CanvasEx.DrawBrushLines(oDIB32, ptPointerF(), [icoXORAlpha])
                        Call oDIB32.Stretch32(.hDC, 0, 0, .Width, .Height)
                      Else
                        Call Me.CanvasEx.DrawBrushLines(m_oCanvasBuffer, ptPointerF(), [icoXORSolid])
                    End If
                
                Case [icShape]
                
                    If (m_lpoIcon.BPP(m_ImageIdx) = [ARGB_Color]) Then
                        Call oDIB32.Create(.Width, .Height, [32_bpp])
                        Call oDIB32.Reset
                        Call Me.CanvasEx.DrawShape(oDIB32, ptPointerF(0).x, ptPointerF(0).y, ptPointerF(1).x, ptPointerF(1).y, [icoXORAlpha])
                        Call oDIB32.Stretch32(.hDC, 0, 0, .Width, .Height)
                      Else
                        Call Me.CanvasEx.DrawShape(m_oCanvasBuffer, ptPointerF(0).x, ptPointerF(0).y, ptPointerF(1).x, ptPointerF(1).y, [icoXORSolid])
                    End If
                    
                Case [icHotSpot]
                
                    Static lHotSpotColor As Long
                    lHotSpotColor = lHotSpotColor Xor &HFFFFFF
                    Call SetPixelV(.hDC, m_lpoIcon.HotSpotX(m_ImageIdx), m_lpoIcon.HotSpotY(m_ImageIdx), lHotSpotColor)
                    Call SetPixelV(.hDC, ptPointerL(0).x, ptPointerL(0).y, lHotSpotColor Xor &HFFFFFF)
            End Select
        End With
    End If
End Sub

Private Sub pvDrawDotGrid()
  
  Dim m_Rect      As RECT2
  Dim hOldBrush   As Long
  Dim lOldBkColor As Long
  Dim hOldPen     As Long
  Dim hPen        As Long
  Dim lColor      As Long
  Dim lIdx        As Long
  
  Dim W As Long
  Dim H As Long
    
    With picCanvas
        
        W = .Width - EDGE3D \ 2 - 1
        H = .Height - EDGE3D \ 2 - 1
        
        If (m_ShowGrid) Then
           
            '-- Set back color and select dot brush
            lOldBkColor = SetBkColor(.hDC, &HC0C0C0)
            hOldBrush = SelectObject(.hDC, m_hDotBrush)
            
            '-- Horizontal dot-lines
            Call BeginPath(.hDC)
                For lIdx = EDGE3D \ 2 To H Step m_ScaleFactor
                    Call pvDrawLine(.hDC, EDGE3D \ 2, lIdx, W, lIdx)
                Next lIdx
            Call EndPath(.hDC)
            Call WidenPath(.hDC)
            Call FillPath(.hDC)
            
            '-- Vertical dot-lines
            Call BeginPath(.hDC)
                For lIdx = EDGE3D \ 2 To W Step m_ScaleFactor
                    Call pvDrawLine(.hDC, lIdx, EDGE3D \ 2, lIdx, H)
                Next lIdx
            Call EndPath(.hDC)
            Call WidenPath(.hDC)
            Call FillPath(.hDC)
            
            '-- Create pen
            Call OleTranslateColor(vb3DShadow, 0, lColor)
            hPen = CreatePen(PS_SOLID, 1, lColor)
            hOldPen = SelectObject(.hDC, hPen)
            
            '-- Restore back color, brush, and destroy pen
            Call SetBkColor(.hDC, lOldBkColor)
            Call SelectObject(.hDC, hOldBrush)
            Call SelectObject(.hDC, hOldPen)
            Call DeleteObject(hPen)
        End If
        
        '-- Erase edge
        hPen = CreatePen(PS_SOLID, 1, &HC0C0C0)
        hOldPen = SelectObject(.hDC, hPen)
        Call pvDrawRect(.hDC, 1, 1, W + 1, H + 1)
        Call pvDrawLine(.hDC, W - m_ShowGrid, 1, W - m_ShowGrid, H - m_ShowGrid + 1)
        Call pvDrawLine(.hDC, 1, H - m_ShowGrid, W - m_ShowGrid + 1, H - m_ShowGrid)
        Call SelectObject(.hDC, hOldPen)
        Call DeleteObject(hPen)
        
        '-- 'Hide' edge
        m_Rect.x2 = .ScaleWidth
        m_Rect.y2 = .ScaleHeight
        Call DrawEdge(.hDC, m_Rect, 4, BF_RECT)
    End With
End Sub

Private Sub pvDrawLine(ByVal lhDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

  Dim uPt As POINTAPI
    
    Call MoveToEx(lhDC, x1, y1, uPt)
    Call LineTo(lhDC, x2, y2)
End Sub

Private Sub pvDrawRect(ByVal lhDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

  Dim uRct      As RECT2
  Dim lTmp      As Long
  Dim uLogBrush As LOGBRUSH
  Dim hOldBrush As Long
  Dim hBrush    As Long
    
    With uLogBrush
        .lbStyle = BS_NULL
        .lbColor = 0
        .lbHatch = 0
    End With
    
    If (x1 = x2 Or y1 = y2) Then
        Call pvDrawLine(lhDC, x1, y1, x2, y2)
        Call SetPixelV(lhDC, x2, y2, 0)
      Else
        hBrush = CreateBrushIndirect(uLogBrush)
        hOldBrush = SelectObject(lhDC, hBrush)
        If (x2 < x1) Then lTmp = x2: x2 = x1 + 1: x1 = lTmp Else x2 = x2 + 1
        If (y2 < y1) Then lTmp = y2: y2 = y1 + 1: y1 = lTmp Else y2 = y2 + 1
        Call Rectangle(lhDC, x1, y1, x2, y2)
        Call SelectObject(lhDC, hOldBrush)
        Call DeleteObject(hBrush)
    End If
End Sub

'//

Private Function pvDIBx(ByVal xCanvas As Single) As Single
    
    pvDIBx = (xCanvas - EDGE3D \ 2) / m_ScaleFactor
    If (m_lpoIcon.BPP(m_ImageIdx) <> [ARGB_Color]) Then
        pvDIBx = Int(pvDIBx)
    End If
End Function

Private Function pvDIBy(ByVal yCanvas As Single) As Single
    
    pvDIBy = (yCanvas - EDGE3D \ 2) / m_ScaleFactor
    If (m_lpoIcon.BPP(m_ImageIdx) <> [ARGB_Color]) Then
        pvDIBy = Int(pvDIBy)
    End If
End Function

'//

'========================================================================================
' Properties
'========================================================================================

Public Property Get Tool() As icToolKeyCts
    Tool = m_Tool
End Property
Public Property Let Tool(ByVal New_Tool As icToolKeyCts)

  Dim sToolID As String
    
    '-- Set new tool index
    m_Tool = New_Tool
    
    '-- Change pointer
    Select Case m_Tool
        Case [icSelectionFrame]: sToolID = "CURSOR_FRAMESELECTION"
        Case [icPencil]:         sToolID = "CURSOR_PENCIL"
        Case [icFloodFill]:      sToolID = "CURSOR_FLOODFILL"
        Case [icColorSelector]:  sToolID = "CURSOR_COLORSELECTOR"
        Case Else:               sToolID = "CURSOR_POINTER"
    End Select
    picCanvas.MouseIcon = LoadResPicture(sToolID, vbResCursor)
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property

Public Property Get ShowGrid() As Boolean
    ShowGrid = m_ShowGrid
End Property
Public Property Let ShowGrid(ByVal New_ShowGrid As Boolean)
    m_ShowGrid = New_ShowGrid
End Property

'//

Public Property Get oFrameSelectionXOR() As cDIB
    If (m_Tool = [icSelectionFrame] And Not m_oDstRect.IsEmpty) Then
        Set oFrameSelectionXOR = m_oXORBuffer
    End If
End Property

Public Property Get oFrameSelectionAND() As cDIB
    If (m_Tool = [icSelectionFrame] And Not m_oDstRect.IsEmpty) Then
        Set oFrameSelectionAND = m_oANDBuffer
    End If
End Property

'//

Public Property Get IsPrivateClipboardAvailable() As Boolean
    IsPrivateClipboardAvailable = (m_oXORClip.hDIB <> 0)
End Property

Public Property Get IsPaletteBlackEntryAvailable(ByVal Index As Byte) As Boolean
    If (Not m_lpoIcon Is Nothing) Then
        With m_lpoIcon
            IsPaletteBlackEntryAvailable = Me.CanvasEx.IsPaletteBlackEntryAvailable(.oXORDIB(m_ImageIdx), .oANDDIB(m_ImageIdx), Index)
        End With
    End If
End Property

'*

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'//

Private Sub UserControl_InitProperties()
    m_ShowGrid = m_def_ShowGrid
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_ShowGrid = PropBag.ReadProperty("ShowGrid", m_def_ShowGrid)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ShowGrid", m_ShowGrid, m_def_ShowGrid)
End Sub
