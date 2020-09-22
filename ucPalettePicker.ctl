VERSION 5.00
Begin VB.UserControl ucPalettePicker 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1740
   ClipControls    =   0   'False
   ForeColor       =   &H80000010&
   MousePointer    =   99  'Custom
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   116
   Begin VB.Timer tmrCursor 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Shape shpEntry 
      BorderColor     =   &H00000000&
      Height          =   225
      Left            =   495
      Top             =   105
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "ucPalettePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucPalettePicker.ctl
' Author:        Carles P.V.
' Dependencies:
' Last revision: 2003.03.28
'================================================

Option Explicit

'-- API:

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BF_RECT         As Long = &HF

Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT2, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)

'-- Private 'Constants':
Private CC_WIDTH  As Long
Private CC_HEIGHT As Long

'-- Private Variables:
Private m_rctControl   As RECT2
Private m_rctCell(255) As RECT2
Private m_aPalette()   As Byte
Private m_hBrush       As Long
Private m_aIdx         As Byte
Private m_R            As Byte
Private m_G            As Byte
Private m_B            As Byte
Private m_LastButton   As Integer

'-- Event Declarations:
Public Event ColorASelected(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal Index As Byte)
Public Event ColorBSelected(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal Index As Byte)
Public Event ColorOver(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal Index As Byte)
Public Event ColorDblClick(Button As Integer, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal Index As Byte)
Public Event MouseOut()



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()
 
  Dim hBitmap        As Long
  Dim aBytes(1 To 8) As Integer
    
    '-- Brush pattern (8x8)
    aBytes(1) = &HAA
    aBytes(2) = &H55
    aBytes(3) = &HAA
    aBytes(4) = &H55
    aBytes(5) = &HAA
    aBytes(6) = &H55
    aBytes(7) = &HAA
    aBytes(8) = &H55
    
    '-- Create brush
    hBitmap = CreateBitmap(8, 8, 1, 1, aBytes(1))
    m_hBrush = CreatePatternBrush(hBitmap)
    Call DeleteObject(hBitmap)
    
    '-- Initialize palette
    ReDim m_aPalette(0)
    
    '-- Initialize pointers
    UserControl.MouseIcon = LoadResPicture("CURSOR_COLORSELECTOR", vbResCursor)
End Sub

Private Sub UserControl_Show()
    
    tmrCursor.Enabled = Ambient.UserMode
End Sub

Private Sub UserControl_Terminate()
    
    '-- Delete pattern brush
    Call DeleteObject(m_hBrush)
End Sub

Private Sub UserControl_Resize()
  
  Dim lEnt As Long
  Dim lx   As Long
  Dim ly   As Long
    
    CC_WIDTH = 11
    CC_HEIGHT = ScaleHeight \ 32: If (CC_HEIGHT < 5) Then CC_HEIGHT = 5
    
    Call UserControl.Size((8 * CC_WIDTH) * Screen.TwipsPerPixelX, (32 * CC_HEIGHT) * Screen.TwipsPerPixelY)
    
    '-- Define rects.
    Call SetRect(m_rctControl, 0, 0, ScaleWidth, ScaleHeight)
    For ly = 0 To 31
        For lx = 0 To 7
            Call SetRect(m_rctCell(lEnt), lx * CC_WIDTH, ly * CC_HEIGHT, lx * CC_WIDTH + CC_WIDTH, ly * CC_HEIGHT + CC_HEIGHT)
            lEnt = lEnt + 1
        Next lx
    Next ly
    
    '-- Resize palette cell cursor and refresh
    Call shpEntry.Move(0, 0, CC_WIDTH - 1, CC_HEIGHT - 1)
    Call Me.Refresh
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub SetPalette(Palette() As Byte)
    
    '-- Assign/set palette
    ReDim m_aPalette(UBound(Palette()) - LBound(Palette()) + 1)
    Call CopyMemory(m_aPalette(0), Palette(0), UBound(Palette()) - LBound(Palette()) + 1)
End Sub

Public Sub GetPalette(Palette() As Byte)
    
    '-- Assign/get palette
    ReDim Palette(UBound(m_aPalette()) - LBound(m_aPalette()) + 1)
    Call CopyMemory(Palette(0), m_aPalette(0), UBound(m_aPalette()) - LBound(m_aPalette()) + 1)
End Sub

Public Sub SetCursor(ByVal Index As Byte, ByVal Show As Boolean)
    
    '-- Set entry visibility and position
    shpEntry.Visible = Show
    shpEntry.BorderColor = &HFFFFFF
    Call shpEntry.Move((Index Mod 8) * CC_WIDTH, (Index \ 8) * CC_HEIGHT)
End Sub

Public Sub Refresh()
    
    '-- Repaint control
    Call pvPaintPalette
    Call UserControl.Refresh
End Sub

'========================================================================================
' Events
'========================================================================================

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (PtInRect(m_rctControl, x, y)) Then

        m_aIdx = (x \ CC_WIDTH) + 8 * (y \ CC_HEIGHT)
        
        If (m_aIdx < (UBound(m_aPalette()) + 1) \ 4) Then
  
            m_R = m_aPalette(4 * m_aIdx + 2)
            m_G = m_aPalette(4 * m_aIdx + 1)
            m_B = m_aPalette(4 * m_aIdx + 0)
            
            RaiseEvent ColorOver(m_R, m_G, m_B, m_aIdx)
          
          Else
            Call ReleaseCapture
            RaiseEvent MouseOut
        End If
    End If
    
    If (GetCapture <> UserControl.hWnd) Then
        Call SetCapture(UserControl.hWnd)
    End If
    If (x < 0 Or y < 0 Or x >= UserControl.ScaleWidth Or y >= UserControl.ScaleHeight) Then
        Call ReleaseCapture
        RaiseEvent MouseOut
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (PtInRect(m_rctControl, x, y)) Then
    
        m_aIdx = (x \ CC_WIDTH) + 8 * (y \ CC_HEIGHT)
        
        If (m_aIdx < (UBound(m_aPalette()) + 1) \ 4) Then
            
            m_R = m_aPalette(4 * m_aIdx + 2)
            m_G = m_aPalette(4 * m_aIdx + 1)
            m_B = m_aPalette(4 * m_aIdx + 0)
            
            Select Case Button
                Case vbLeftButton
                    RaiseEvent ColorASelected(m_R, m_G, m_B, m_aIdx)
                Case vbRightButton
                    RaiseEvent ColorBSelected(m_R, m_G, m_B, m_aIdx)
            End Select
        End If
    End If
    
    m_LastButton = Button
End Sub

Private Sub UserControl_DblClick()

    If (m_aIdx < (UBound(m_aPalette()) + 1) \ 4) Then
        If (m_LastButton = vbLeftButton Or m_LastButton = vbRightButton) Then
            RaiseEvent ColorDblClick(m_LastButton, m_R, m_G, m_B, m_aIdx)
        End If
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvPaintPalette()
    
  Dim lEnt  As Long
  Dim lEnts As Long
  Dim lClr  As Long
  Dim lx    As Long
  Dim ly    As Long
  
    If (UBound(m_aPalette())) Then
    
        lEnts = (UBound(m_aPalette()) + 1) \ 4
        
        '-- Paint palette sample cells
        Do
            If (lEnt < lEnts) Then
                lClr = RGB(m_aPalette(4 * lEnt + 2), m_aPalette(4 * lEnt + 1), m_aPalette(4 * lEnt + 0))
              Else
                lClr = -1
            End If
            Call pvDrawRectangle(hDC, m_rctCell(lEnt), lClr)
            Call DrawEdge(hDC, m_rctCell(lEnt), BDR_SUNKENOUTER, BF_RECT)

            lx = lx + 1: If (lx = 8) Then lx = 0: ly = ly + 1
            lEnt = lEnt + 1
        Loop Until lEnt = 256
    End If
End Sub

Private Function pvDrawRectangle(ByVal hDC As Long, lpRect As RECT2, ByVal lColor As Long)

  Dim hBrush As Long
  
    If (lColor > -1) Then
        hBrush = CreateSolidBrush(lColor)
        Call FillRect(hDC, lpRect, hBrush)
        Call DeleteObject(hBrush)
      Else
        Call FillRect(hDC, lpRect, m_hBrush)
    End If
End Function

Private Sub tmrCursor_Timer()
    shpEntry.BorderColor = shpEntry.BorderColor Xor &HFFFFFF
End Sub
