VERSION 5.00
Begin VB.Form fCapture 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Screen Capture"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3525
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   660
      Width           =   1050
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Capture area"
      Height          =   1875
      Left            =   165
      TabIndex        =   8
      Top             =   120
      Width           =   3210
      Begin VB.OptionButton optArea 
         Caption         =   "&16 x 16"
         Height          =   225
         Index           =   0
         Left            =   195
         TabIndex        =   0
         Top             =   375
         Width           =   1050
      End
      Begin VB.OptionButton optArea 
         Caption         =   "&24 x 24"
         Height          =   225
         Index           =   1
         Left            =   195
         TabIndex        =   1
         Top             =   705
         Width           =   1050
      End
      Begin VB.OptionButton optArea 
         Caption         =   "&32 x 32"
         Height          =   225
         Index           =   2
         Left            =   195
         TabIndex        =   2
         Top             =   1035
         Width           =   1050
      End
      Begin VB.OptionButton optArea 
         Caption         =   "&48 x 48"
         Height          =   225
         Index           =   3
         Left            =   195
         TabIndex        =   3
         Top             =   1365
         Width           =   1095
      End
      Begin VB.OptionButton optArea 
         Caption         =   "&Drag custom"
         Height          =   225
         Index           =   4
         Left            =   1395
         TabIndex        =   4
         Top             =   375
         Width           =   1455
      End
      Begin VB.OptionButton optArea 
         Caption         =   "&Custom"
         Height          =   225
         Index           =   5
         Left            =   1395
         TabIndex        =   5
         Top             =   705
         Width           =   1320
      End
      Begin iICO.ucUpDownBox ucWidth 
         Height          =   315
         Left            =   2190
         TabIndex        =   11
         Top             =   1005
         Width           =   825
         _ExtentX        =   1429
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin iICO.ucUpDownBox ucHeight 
         Height          =   315
         Left            =   2190
         TabIndex        =   12
         Top             =   1365
         Width           =   825
         _ExtentX        =   1482
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblWidth 
         Caption         =   "Width"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1665
         TabIndex        =   10
         Top             =   1065
         Width           =   615
      End
      Begin VB.Label lblHeight 
         Caption         =   "Height"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1665
         TabIndex        =   9
         Top             =   1425
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCapture 
      Caption         =   "Capture!"
      Default         =   -1  'True
      Height          =   390
      Left            =   3525
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   210
      Width           =   1050
   End
   Begin VB.Menu mnuCaptureTop 
      Caption         =   "Capture"
      Visible         =   0   'False
      Begin VB.Menu mnuCapture 
         Caption         =   "Show/Hide dialog"
         Index           =   0
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "Show message"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "Hide after capture"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "About..."
         Index           =   5
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "Exit"
         Index           =   6
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuCapture 
         Caption         =   "Cancel"
         Index           =   8
      End
   End
End
Attribute VB_Name = "fCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- API

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

Private Const RDW_INVALIDATE       As Long = &H1
Private Const RDW_FRAME            As Long = &H400
Private Const RDW_UPDATENOW        As Long = &H100
Private Const R2_NOT               As Long = 6
Private Const SRCCOPY              As Long = &HCC0020
Private Const BS_NULL              As Long = 1
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

'//

Private mouse_raised As Boolean  ' mouse_event called
Private mouse_downed As Boolean  ' mouse 'manualy downed'

'//

Private m_dskhWnd    As Long     ' Desktop window
Private m_dskhDC     As Long     ' Desktop DC
Private m_dskRect    As RECT2    ' Desktop 'work' area

Private m_bCustom    As Boolean  ' Custom area flag
Private m_ptStart    As POINTAPI ' Start point
Private m_ptCurrent  As POINTAPI ' Current point
Private m_ptLast     As POINTAPI ' Last point (Invert restoration)
Private m_clpW       As Long     ' Selection (clipboard) width
Private m_clpH       As Long     ' Selection (clipboard) height

Private m_bTrapping  As Boolean  ' Start capture (Mouse events flag)
Private m_bDone      As Boolean  ' Capture done (Mouse up)

'//

Private Sub Form_Load()

    Set Me.Icon = Nothing
    
    Call mMisc.RemoveButtonBorderEnhance(Me.cmdCapture)
    Call mMisc.RemoveButtonBorderEnhance(Me.cmdCancel)
    
    With ucWidth
        .Min = 1
        .Max = Screen.Width \ Screen.TwipsPerPixelX
        .Value = 16
    End With
    With ucHeight
        .Min = 1
        .Max = Screen.Height \ Screen.TwipsPerPixelY
        .Value = 16
    End With
End Sub

'==========================================================================================
' Capture area
'==========================================================================================

Private Sub optArea_Click(Index As Integer)
  
  Dim bDragCustom As Boolean
    
    m_bCustom = False
    
    Select Case Index
        Case 0 '-- 16 x 16
            m_clpW = 16
            m_clpH = 16
        Case 1 '-- 24 x 24
            m_clpW = 24
            m_clpH = 24
        Case 2 '-- 32 x 32
            m_clpW = 32
            m_clpH = 32
        Case 3 '-- 48 x 48
            m_clpW = 48
            m_clpH = 48
        Case 4 '-- Drag custom
            m_bCustom = True
        Case 5 '-- Custom
            m_clpW = ucWidth.Value
            m_clpH = ucHeight.Value
    End Select
      
    '-- Enable/Disable <Custom> text boxes
    bDragCustom = optArea(5).Value
    
    lblWidth.Enabled = bDragCustom
    ucWidth.Enabled = bDragCustom
    
    lblHeight.Enabled = bDragCustom
    ucHeight.Enabled = bDragCustom
    
    If (bDragCustom) Then
        On Error Resume Next
        Call ucWidth.SetFocus
        On Error GoTo 0
    End If
End Sub

'==========================================================================================
' Custom capture
'==========================================================================================

Private Sub ucWidth_Change()
    
    '-- Update
    m_clpW = ucWidth.Value
End Sub

Private Sub ucHeight_Change()
    
    '-- Update
    m_clpH = ucHeight.Value
End Sub

'==========================================================================================
' Quit/Hide dialog
'==========================================================================================
    
Private Sub cmdCancel_Click()

    fMain.Enabled = True
    Call fCapture.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    '-- Close button pressed
    If (UnloadMode = vbFormControlMenu) Then
        Call cmdCancel_Click
    End If
End Sub

'==========================================================================================
' Capture!
'==========================================================================================

Private Sub cmdCapture_Click()

 Dim dsk_Status As Long
 Dim hBrush     As Long
 Dim uLogBrush  As LOGBRUSH
 
 Dim Me_Rect    As RECT2
 Dim bSuccess   As Boolean
 
    '-- Enable mouse
    m_bTrapping = True
    
    '-- Get desktop DC
    m_dskhWnd = GetDesktopWindow
    m_dskhDC = GetWindowDC(m_dskhWnd)
    
    '-- Save desktop DC status
    dsk_Status = SaveDC(m_dskhDC)
    
    '-- Prepare desktop DC...
    '-- Set draw mode: NOT
    Call SetROP2(m_dskhDC, R2_NOT)
    '-- Create a null brush
    uLogBrush.lbStyle = BS_NULL
    hBrush = CreateBrushIndirect(uLogBrush)
    Call SelectObject(m_dskhDC, hBrush)
    
    '-- Get cursor position and hide it
    Call GetCursorPos(m_ptStart)
    Call GetCursorPos(m_ptLast)
    Call GetCursorPos(m_ptCurrent)
    Call ShowCursor(0)
    
    '-- Pass mouse control to form:
    Call SetCapture(Me.hWnd)
    mouse_downed = False
    mouse_raised = False
    Call GetWindowRect(Me.hWnd, Me_Rect)
    Call SetCursorPos(Me_Rect.x1, Me_Rect.y1)
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    Call SetCursorPos(m_ptStart.x, m_ptStart.y)

    '-- Hide Main form
    fMain.WindowState = vbMinimized
    DoEvents
    '-- Redraw rectangle (force paint) behind form
    Call RedrawWindow(WindowFromPoint(m_ptStart.x, m_ptStart.y), Me_Rect, ByVal 0, RDW_FRAME Or RDW_INVALIDATE Or RDW_UPDATENOW)
    '-- Paint first selection pointer/rectangle
    If (m_bCustom) Then
        Call pvDrawPointer(m_dskhDC, m_ptLast.x, m_ptLast.y)
      Else
        Call Rectangle(m_dskhDC, m_ptLast.x, m_ptLast.y, m_ptLast.x + m_clpW, m_ptLast.y + m_clpH)
    End If
    
    '-- Start...
    m_bDone = False
    Do: Call pvCheckMouse
        DoEvents
    Loop Until m_bDone

    '-- Copy selected area to Clipboard
    bSuccess = pvCopyToClipboard
    
    '-- Restore cursor position and show it
    Call SetCursorPos(m_ptCurrent.x, m_ptCurrent.y)
    Call ShowCursor(1)
    
    '-- A little message for user
    If (bSuccess) Then
        Call MsgBox("Bitmap (" & m_clpW & "x" & m_clpH & " pixels) successfully captured to Clipboard.", vbInformation)
      Else
        Call MsgBox("Unexpexted error capturing area to Clipboard.", vbExclamation)
    End If
    
    '-- Restore/Free DC
    Call RestoreDC(m_dskhDC, dsk_Status)
    Call ReleaseDC(m_dskhWnd, m_dskhDC)
    
    '-- All done
    m_bTrapping = False
    fMain.Enabled = True
    fMain.WindowState = vbNormal
    Call Me.Hide
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (mouse_raised = False) Then
    
        '-- Get start point
        Call GetCursorPos(m_ptStart)
        
        If (Button = vbLeftButton And m_bTrapping) Then
            
            '-- Mouse pressed flag
            If (mouse_downed) Then
                mouse_raised = True
            End If
            mouse_downed = True
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (Button = vbLeftButton And m_bTrapping And mouse_downed) Then
        
        If (m_bCustom) Then
            '-- Restore last pointer
            Call pvDrawPointer(m_dskhDC, m_ptCurrent.x, m_ptCurrent.y)
            '-- Restore last painted rectangle
            Call Rectangle(m_dskhDC, m_ptStart.x, m_ptStart.y, m_ptCurrent.x, m_ptCurrent.y)
          Else
            '-- Restore last painted rectangle
            Call Rectangle(m_dskhDC, m_ptCurrent.x, m_ptCurrent.y, m_ptCurrent.x + m_clpW, m_ptCurrent.y + m_clpH)
        End If
        
        '-- Areas sensible to mouse movements (for example, toolbars, menus, etc) are updated
        '   before last rectangle restoration, so this rectangle is in part painted again
        '   instead of being erased.
        '   Cursor position will be restored after clipboard copy.
        Call GetWindowRect(GetDesktopWindow, m_dskRect)
        Call SetCursorPos(m_dskRect.x2, m_dskRect.y2)

        '-- Capture done
        Call ReleaseCapture
        m_bDone = True
    End If
End Sub

Private Sub Form_DblClick()

    '-- 'Preserve mouse traping'
    mouse_raised = True
End Sub

'//

Private Sub pvCheckMouse()

    Call GetCursorPos(m_ptCurrent)
    
    If (m_ptLast.x <> m_ptCurrent.x Or m_ptLast.y <> m_ptCurrent.y) Then
        
        If (m_bCustom) Then
        
            '-- Restore last pointer
            Call pvDrawPointer(m_dskhDC, m_ptLast.x, m_ptLast.y)
            '-- Draw current pointer
            Call pvDrawPointer(m_dskhDC, m_ptCurrent.x, m_ptCurrent.y)
            
            If (mouse_raised) Then
                '-- Restore last rectangle
                Call Rectangle(m_dskhDC, m_ptStart.x, m_ptStart.y, m_ptLast.x, m_ptLast.y)
                '-- Paint current rectangle
                Call Rectangle(m_dskhDC, m_ptStart.x, m_ptStart.y, m_ptCurrent.x, m_ptCurrent.y)
            End If
      
          Else
            '-- Restore last rectangle
            Call Rectangle(m_dskhDC, m_ptLast.x, m_ptLast.y, m_ptLast.x + m_clpW, m_ptLast.y + m_clpH)
            '-- Paint current rectangle
            Call Rectangle(m_dskhDC, m_ptCurrent.x, m_ptCurrent.y, m_ptCurrent.x + m_clpW, m_ptCurrent.y + m_clpH)
        End If
        
        m_ptLast = m_ptCurrent
        Call SetCapture(Me.hWnd)
    End If
End Sub

Private Sub pvDrawPointer(ByVal hDestDC As Long, ByVal x1 As Long, ByVal y1 As Long)
        
  Dim uPt As POINTAPI
    
    '-- Draw a little pointer
    Call MoveToEx(hDestDC, x1 - 5, y1, uPt)
    Call LineTo(hDestDC, x1 + 6, y1)
    Call MoveToEx(hDestDC, x1, y1 - 5, uPt)
    Call LineTo(hDestDC, x1, y1 + 6)
End Sub

'==========================================================================================
' Clipboard
'==========================================================================================

Private Function pvCopyToClipboard() As Boolean

  Dim oDIB32 As New cDIB
     
    '-- Get selection dimensions
    If (m_bCustom) Then
        If (m_ptCurrent.x < m_ptStart.x) Then m_clpW = m_ptStart.x - m_ptCurrent.x + 1: m_ptStart.x = m_ptCurrent.x Else m_clpW = m_ptCurrent.x - m_ptStart.x + 1
        If (m_ptCurrent.y < m_ptStart.y) Then m_clpH = m_ptStart.y - m_ptCurrent.y + 1: m_ptStart.y = m_ptCurrent.y Else m_clpH = m_ptCurrent.y - m_ptStart.y + 1
    End If
    
    '-- Is there something selected ?
    If (m_clpW > 0 And m_clpH > 0) Then
        Call oDIB32.Create(m_clpW, m_clpH, [32_bpp])
        Call oDIB32.LoadBlt(m_dskhDC, m_ptStart.x, m_ptStart.y, m_clpW, m_clpH)
        pvCopyToClipboard = CBool(oDIB32.CopyToClipboard)
    End If
End Function

