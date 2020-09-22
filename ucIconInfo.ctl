VERSION 5.00
Begin VB.UserControl ucIconInfo 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   ClipControls    =   0   'False
   ScaleHeight     =   14
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
End
Attribute VB_Name = "ucIconInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucInfo.ctl
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

Private Const BDR_RAISEDINNER     As Long = &H4
Private Const BDR_SUNKENOUTER     As Long = &H2
Private Const BF_RECT             As Long = &HF

Private Const COLOR_BTNFACE       As Long = 15

Private Const DFC_SCROLL          As Long = &H3
Private Const DFCS_SCROLLSIZEGRIP As Long = &H8

Private Const WM_NCLBUTTONDOWN    As Long = &HA1
Private Const HTBOTTOMRIGHT       As Long = &H11

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT2, lpSourceRect As RECT2) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT2, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hDC As Long, ByVal pszPath As String, ByVal dx As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
                         
'//

'-- Property Variables:
Private m_TextFile         As String
Private m_TextInfo1        As String
Private m_TextInfo2        As String
Private m_TextInfo3        As String

'-- Private Variables:
Private m_BarRect          As RECT2
Private m_SizeGripRect     As RECT2
Private m_EdgeRect(1 To 4) As RECT2
Private m_TextRect(1 To 4) As RECT2



'========================================================================================
' UserControl
'========================================================================================

Public Sub Refresh()

  Dim lEdg       As Long
  Dim sCmpctPath As String
    
    '-- Erase background
    Call FillRect(hDC, m_BarRect, GetSysColorBrush(COLOR_BTNFACE))
    '-- Draw edges
    For lEdg = 1 To 4
        Call DrawEdge(hDC, m_EdgeRect(lEdg), BDR_SUNKENOUTER, BF_RECT)
    Next lEdg
    '-- Draw size grip
    Call DrawFrameControl(hDC, m_SizeGripRect, DFC_SCROLL, DFCS_SCROLLSIZEGRIP)
    '-- Draw text
    sCmpctPath = pvCompactPath(m_TextFile, pvTextFileWidth)
    Call DrawText(hDC, sCmpctPath, Len(sCmpctPath), m_TextRect(1), &H0)
    Call DrawText(hDC, m_TextInfo1, Len(m_TextInfo1), m_TextRect(2), &H0)
    Call DrawText(hDC, m_TextInfo2, Len(m_TextInfo2), m_TextRect(3), &H0)
    Call DrawText(hDC, m_TextInfo3, Len(m_TextInfo3), m_TextRect(4), &H0)
   
    Call UserControl.Refresh
End Sub

Private Sub UserControl_Resize()
    
  Const INFO_WIDTH1 As Long = 150
  Const INFO_WIDTH2 As Long = 100
  Const INFO_WIDTH3 As Long = 155
   
  Dim W    As Long
  Dim H    As Long
  Dim SG   As Long
  Dim lPnl As Long
  
    W = ScaleWidth
    H = 18
    UserControl.Height = 18 * Screen.TwipsPerPixelY
    
    On Error Resume Next
    
    '-- Check parent form window state:
    '   Size Grip (Show/Hide)
    If (Parent.WindowState = vbMaximized) Then
        SG = 0
      Else
        SG = H
    End If
    
    '-- Set main Rect. and size grip Rect.
    Call SetRect(m_BarRect, 0, 0, W, H)
    Call SetRect(m_SizeGripRect, W - SG, 0, W, H)
    
    '-- Set text Rects. (Edge and text)
    Call SetRect(m_EdgeRect(1), 0, 0, W - INFO_WIDTH1 - INFO_WIDTH2 - INFO_WIDTH3 - SG - 6, H)
    Call SetRect(m_EdgeRect(2), W - INFO_WIDTH1 - INFO_WIDTH2 - INFO_WIDTH3 - SG - 4, 0, W - SG - INFO_WIDTH2 - INFO_WIDTH3 - 4, H)
    Call SetRect(m_EdgeRect(3), W - INFO_WIDTH2 - INFO_WIDTH3 - SG - 2, 0, W - INFO_WIDTH3 - SG - 2, H)
    Call SetRect(m_EdgeRect(4), W - INFO_WIDTH3 - SG, 0, W - SG, H)
    For lPnl = 1 To 4
        Call CopyRect(m_TextRect(lPnl), m_EdgeRect(lPnl))
        With m_TextRect(lPnl)
            .x1 = .x1 + 4
            .y1 = .y1 + 2
            .x2 = .x2 - 4
        End With
    Next lPnl
    
    '-- Refresh
    Call Me.Refresh
    
    On Error GoTo 0
End Sub

'//

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = vbLeftButton And x > m_SizeGripRect.x1) Then
        Call ReleaseCapture
        Call SendMessage(Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (x > m_SizeGripRect.x1) Then
        UserControl.MousePointer = vbSizeNWSE
      Else
        UserControl.MousePointer = vbDefault
   End If
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Let TextFile(ByVal New_TextFile As String)
    m_TextFile = New_TextFile
End Property
Public Property Get TextFile() As String
    TextFile = m_TextFile
End Property

Public Property Let TextInfo1(ByVal New_TextInfo1 As String)
    m_TextInfo1 = New_TextInfo1
End Property
Public Property Get TextInfo1() As String
    TextInfo1 = m_TextInfo1
End Property

Public Property Let TextInfo2(ByVal New_TextInfo2 As String)
    m_TextInfo2 = New_TextInfo2
End Property
Public Property Get TextInfo2() As String
    TextInfo2 = m_TextInfo2
End Property

Public Property Let TextInfo3(ByVal New_TextInfo3 As String)
    m_TextInfo3 = New_TextInfo3
End Property
Public Property Get TextInfo3() As String
    TextInfo3 = m_TextInfo3
End Property

'//

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set UserControl.Font = Ambient.Font
End Sub

'========================================================================================
' Private
'========================================================================================

Public Function pvTextFileWidth() As Long
    pvTextFileWidth = m_TextRect(1).x2 - m_TextRect(1).x1
End Function

Private Function pvCompactPath(ByVal FullPath As String, ByVal Width As Long) As String

  Dim lZeroPos As Long

    '-- Compact
    Call PathCompactPath(UserControl.hDC, ByVal FullPath, Width)

    '-- Remove all trailing Chr$(0)'s
    lZeroPos = InStr(1, FullPath, Chr$(0))
    If (lZeroPos > 0) Then
        pvCompactPath = Left$(FullPath, lZeroPos - 1)
      Else
        pvCompactPath = FullPath
    End If
End Function

