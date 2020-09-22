VERSION 5.00
Begin VB.UserControl ucIconList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   FillStyle       =   4  'Upward Diagonal
   KeyPreview      =   -1  'True
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   Begin VB.TextBox txtDummy 
      Height          =   405
      Left            =   -195
      TabIndex        =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.VScrollBar ucBar 
      Height          =   1110
      Left            =   900
      Max             =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "ucIconList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucIconList.ctl
' Author:        Carles P.V.
' Dependencies:  cIcon.cls
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

Private Const SM_CXVSCROLL As Long = &H2
Private Const DT_CENTER    As Long = &H1
Private Const PS_SOLID     As Long = &H0
Private Const BS_NULL      As Long = &H1

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

'//

'-- Private Types:
Private Type tItem
    Text      As String
    ItemData  As Long
End Type

'-- Default Property Values:
Private Const m_def_ColorScreen = vbWindowBackground
Private Const mdef_ThumbnailHeight As Integer = 72
Private Const mdef_ItemHeightPad   As Integer = 20

'-- Property Variables:
Private m_List()          As tItem     ' List array of items
Private m_ColorScreen     As OLE_COLOR ' Thumbnail selection BackColor
Private m_ThumbnailHeight As Integer   ' Max. thumbnail height
Private m_ListIndex       As Integer   ' Current list index

'-- Private Variables:
Private m_lpoIcon         As cIcon     ' Pointer to cIcon object
Private m_LastIndex       As Boolean   ' Last selected item
Private m_MouseDown       As Boolean   ' Mouse down flag
Private m_LastBar         As Integer   ' Last scroll bar value
Private m_VisibleRows     As Integer   ' Visible rows
Private m_PerfectRowPad   As Boolean   ' Visible rows padding
Private m_ControlRect     As RECT2     ' User control rectangle (clearing)
Private m_RectExt()       As RECT2     ' Item rectangle
Private m_RectTxt()       As RECT2     ' Item text rectangle
Private m_RectSel()       As RECT2     ' Item selection rectangle
Private m_RectImg()       As RECT2     ' Item image rectangle
Private m_RectRowPad      As RECT2     ' Last row padding rectangle (background erasing)
Private m_Clr3DDKShadow   As Long      ' 3D dark shadow color
Private m_ClrButtonFace   As Long      ' 3D object color

'-- Event Declarations:
Public Event Click()
Public Event KeyDown(ByVal KeyCode As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Item As Integer)



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()

    '-- Initialize items array
    ReDim m_List(0)
    
    '-- Initialize position flags
    m_LastIndex = -1
    m_LastBar = -1
    
    '-- Set system default scroll bar width
    ucBar.Width = GetSystemMetrics(SM_CXVSCROLL)
End Sub

Private Sub UserControl_Terminate()

    '-- Destroy Icon reference
    If (Not m_lpoIcon Is Nothing) Then
        Set m_lpoIcon = Nothing
    End If
    
    '-- Clear items array
    Erase m_List()
End Sub

'//
    
Private Sub UserControl_Show()
    '-- Refresh control
    Call Me.Refresh
End Sub

Private Sub UserControl_Resize()
    
  Dim sVisibleRows As Single
    
    '-- Check minimum height (one row)
    If (ScaleHeight < m_ThumbnailHeight + mdef_ItemHeightPad) Then
        Height = ((m_ThumbnailHeight + mdef_ItemHeightPad) + (Height \ Screen.TwipsPerPixelY - ScaleHeight)) * Screen.TwipsPerPixelY
    End If
    
    '-- Get visible rows
    sVisibleRows = ScaleHeight / (m_ThumbnailHeight + mdef_ItemHeightPad)
    
    '-- Perfect row adjustment [?]
    If (sVisibleRows = Int(sVisibleRows)) Then
        m_VisibleRows = sVisibleRows
        m_PerfectRowPad = True
      Else
        m_VisibleRows = Int(sVisibleRows) + 1
        m_PerfectRowPad = False
    End If

    '-- Calc. items rects.
    Call pvCalculateRects
    
    '-- Scroll bar
    ucBar.Visible = False
    Call ucBar.Move(ScaleWidth - ucBar.Width, 0, ucBar.Width, ScaleHeight)
    Call pvReadjustBar
    
    '-- Refresh (erase background and refresh whole list)
    Call pvRectangle(hDC, m_ControlRect, m_ClrButtonFace, m_ClrButtonFace)
    Call Me.Refresh
End Sub

'//

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode

        Case vbKeyUp       '-- Row up
            If (m_ListIndex > 0) Then
                Let ListIndex = m_ListIndex - 1
            End If
        
        Case vbKeyDown     '-- Row down
            If (m_ListIndex < UBound(m_List) - 1) Then
                Let ListIndex = m_ListIndex + 1
            End If

        Case vbKeyPageUp   '-- Page up
            If (m_ListIndex > m_VisibleRows) Then
                Let ListIndex = m_ListIndex - m_VisibleRows - (Not m_PerfectRowPad)
              Else
                Let ListIndex = 0
            End If
       
        Case vbKeyPageDown '-- Page down
            If (m_ListIndex < UBound(m_List) - m_VisibleRows - 1) Then
                Let ListIndex = m_ListIndex + m_VisibleRows + (Not m_PerfectRowPad)
              Else
                Let ListIndex = UBound(m_List) - 1
            End If

        Case vbKeyHome     '-- Start
            Let ListIndex = 0

        Case vbKeyEnd      '-- End
            Let ListIndex = UBound(m_List) - 1
    End Select
    
    RaiseEvent KeyDown(KeyCode)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim nItm As Integer
  
    If (UBound(m_List)) Then
        
        '-- Over item [?]
        m_MouseDown = Not -PtInRect(m_RectRowPad, x, y)
        '-- Set item
        nItm = ucBar.Value + (y \ (m_ThumbnailHeight + mdef_ItemHeightPad))
        If (nItm > -1 And nItm < UBound(m_List)) Then
            If (Button = vbLeftButton) Then
                Let ListIndex = nItm
            End If
        End If
        RaiseEvent MouseDown(Button, nItm)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
    If (UBound(m_List)) Then
  
        If (m_MouseDown And Button = vbLeftButton) Then
            '-- Item selected
            RaiseEvent Click
        End If
        m_MouseDown = False
    End If
End Sub

'========================================================================================
' Scroll bar
'========================================================================================

Private Sub ucBar_Change()
    If (m_LastBar <> ucBar.Value) Then
        m_LastBar = ucBar.Value
        Call Me.Refresh
    End If
End Sub

Private Sub ucBar_Scroll()
    Call ucBar_Change
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub SetIconSource(oIcon As cIcon)

    '-- Set reference to source Icon object
    Set m_lpoIcon = oIcon
End Sub

Public Sub AddItem(ByVal Text, Optional ByVal ItemData As Long = 0)

    With m_List(UBound(m_List))
        .Text = CStr(Text)
        .ItemData = ItemData
    End With
    ReDim Preserve m_List(UBound(m_List) + 1)
    
    Call pvReadjustBar
End Sub

Public Sub Clear()

    '-- Reset/Hide scroll bar
    m_LastBar = 0: ucBar.Value = 0
    m_LastBar = -1
    ucBar.Visible = False
    '-- Clean control
    Call pvRectangle(hDC, m_ControlRect, m_ClrButtonFace, m_ClrButtonFace)
    '-- Reset items array / flags
    ReDim m_List(0)
    m_LastIndex = -1
    m_ListIndex = -1
End Sub

Public Sub Refresh()

    '-- Paint visible items
    If (Ambient.UserMode And Extender.Visible) Then
        Call pvDrawList
        Call UserControl.Refresh
    End If
End Sub

Public Sub RefreshItem(ByVal nIndex As Integer)

    '-- Paint single item
    If (Ambient.UserMode And Extender.Visible) Then
        Call pvDrawItem(nIndex)
        Call UserControl.Refresh
    End If
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    ucBar.Enabled = New_Enabled
End Property

Public Property Get ListCount() As Integer
Attribute ListCount.VB_MemberFlags = "400"
    ListCount = UBound(m_List)
End Property

Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = m_ListIndex
End Property
Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    
    '-- Check/Set
    If (New_ListIndex < 0 Or UBound(m_List) = 0) Then
        m_ListIndex = -1
      Else
        m_ListIndex = New_ListIndex
    End If
    m_LastIndex = m_ListIndex

    '-- Ensure visible current selected item
    If (m_ListIndex < ucBar And m_ListIndex > -1) Then
        ucBar.Value = m_ListIndex
      ElseIf (m_ListIndex > ucBar + m_VisibleRows - 1 + (Not m_PerfectRowPad)) Then
        ucBar.Value = m_ListIndex - m_VisibleRows + 1 - (Not m_PerfectRowPad)
      Else
        Call Me.Refresh
    End If
    
    '-- Raise <Click> event [?]
    If (Not m_MouseDown) Then RaiseEvent Click
End Property

Public Property Get TopIndex() As Integer
Attribute TopIndex.VB_MemberFlags = "400"
    TopIndex = ucBar.Value
End Property
Public Property Let TopIndex(ByVal New_TopIndex As Integer)
    '-- Check
    If (New_TopIndex > ucBar.Max) Then
        New_TopIndex = ucBar.Max
    End If
    If (New_TopIndex < 0) Then
        New_TopIndex = 0
    End If
    '-- Set and refresh
    m_LastBar = New_TopIndex
    ucBar.Value = New_TopIndex
    Call Me.Refresh
End Property

Public Property Get VisibleRows() As Integer
    VisibleRows = m_VisibleRows
End Property

Public Property Get PerfectRowPad() As Boolean
    PerfectRowPad = m_PerfectRowPad
End Property

'//

Public Property Get ThumbnailHeight() As Integer
    ThumbnailHeight = m_ThumbnailHeight
End Property
Public Property Let ThumbnailHeight(ByVal New_ThumbnailHeight As Integer)
    m_ThumbnailHeight = New_ThumbnailHeight
    '-- Clean control and refresh
    Call pvRectangle(hDC, m_ControlRect, m_ClrButtonFace, m_ClrButtonFace)
    Call UserControl_Resize
End Property

Public Property Get ColorScreen() As OLE_COLOR
    ColorScreen = m_ColorScreen
End Property
Public Property Let ColorScreen(ByVal New_ColorScreen As OLE_COLOR)
    '-- Translate color and refresh
    Call OleTranslateColor(New_ColorScreen, 0, m_ColorScreen)
    Call Me.Refresh
End Property

'-- Item data...

Public Property Get ItemText(ByVal nIndex As Integer) As Variant
    ItemText = m_List(nIndex).Text
End Property
Public Property Let ItemText(ByVal nIndex As Integer, ByVal New_Text)
    m_List(nIndex).Text = CStr(New_Text)
End Property

Public Property Get ItemData(ByVal nIndex As Integer) As Long
    ItemData = m_List(nIndex).ItemData
End Property
Public Property Let ItemData(ByVal nIndex As Integer, ByVal New_ItemData As Long)
    m_List(nIndex).ItemData = New_ItemData
End Property

'*

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'//

Private Sub UserControl_InitProperties()

    '-- Defaults
    m_ThumbnailHeight = mdef_ThumbnailHeight
    m_ColorScreen = m_def_ColorScreen
    
    '-- Set colors
    Call pvSetColors
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    '-- Set colors
    Call pvSetColors
    '-- Set font
    UserControl.Font = Ambient.Font
    
    '-- Read props.
    With PropBag
        UserControl.Enabled = .ReadProperty("Enabled", True)
        ThumbnailHeight = .ReadProperty("ThumbnailHeight", mdef_ThumbnailHeight)
        m_ColorScreen = .ReadProperty("ColorScreen", m_def_ColorScreen)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("ThumbnailHeight", m_ThumbnailHeight, mdef_ThumbnailHeight)
        Call .WriteProperty("ColorScreen", m_ColorScreen, m_def_ColorScreen)
    End With
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvDrawList()

  Dim nItm As Integer
    
    '-- Draw visible rows
    For nItm = ucBar.Value To ucBar.Value + m_VisibleRows - 1
        Call pvDrawItem(nItm)
    Next nItm
    
    '-- Clear pad row
    If (Not m_PerfectRowPad And ucBar = ucBar.Max) Then
        Call pvRectangle(hDC, m_RectRowPad, m_ClrButtonFace, m_ClrButtonFace)
    End If
End Sub

Private Sub pvDrawItem(ByVal nItm As Integer)
  
  Dim nRct As Integer
    
    '-- Items [?]
    If (nItm < UBound(m_List())) Then
    
        '-- Visible item [?]
        If (nItm >= ucBar.Value And nItm < ucBar.Value + m_VisibleRows) Then
            
            '-- Rect. number
            nRct = nItm - ucBar.Value
            
            '-- Draw background
            Call pvRectangle(hDC, m_RectExt(nRct), m_ClrButtonFace, m_ClrButtonFace)
            If (m_ListIndex = nItm) Then
                Call pvRectangle(hDC, m_RectSel(nRct), m_ColorScreen, m_Clr3DDKShadow)
            End If
            
            '-- Draw text
            Call DrawText(hDC, m_List(nItm).Text, Len(m_List(nItm).Text), m_RectTxt(nRct), DT_CENTER)
            
            '-- Draw thumbnail
            If (Not m_lpoIcon Is Nothing) Then
                Call m_lpoIcon.DrawIconFit(nItm, hDC, m_RectImg(nRct).x1, m_RectImg(nRct).y1, True, m_RectImg(nRct).x2 - m_RectImg(nRct).x1, m_RectImg(nRct).y2 - m_RectImg(nRct).y1)
            End If
        End If
    End If
End Sub

Private Sub pvRectangle(ByVal hDC As Long, m_Rect As RECT2, ByVal FillColor As Long, ByVal BorderColor As Long)

  Dim hBrush    As Long
  Dim hOldBrush As Long
  Dim hPen      As Long
  Dim hOldPen   As Long
    
    '-- Create Pen / Brush
    hPen = CreatePen(PS_SOLID, 1, BorderColor)
    hBrush = CreateSolidBrush(FillColor)
    
    '-- Select into given DC
    hOldPen = SelectObject(hDC, hPen)
    hOldBrush = SelectObject(hDC, hBrush)
    
    '-- Draw rectangle
    Call Rectangle(hDC, m_Rect.x1, m_Rect.y1, m_Rect.x2, m_Rect.y2)
    
    '-- Destroy used objects
    Call SelectObject(hDC, hOldBrush)
    Call DeleteObject(hBrush)
    Call SelectObject(hDC, hOldPen)
    Call DeleteObject(hPen)
End Sub

Private Sub pvReadjustBar()

    On Error Resume Next
    
    If (UBound(m_List) > m_VisibleRows + (Not m_PerfectRowPad)) Then
        If (Not ucBar.Visible) Then
            '-- Show scroll bar
            ucBar.Visible = True
            Call ucBar.Refresh
            Call pvUpdateRectRight(ScaleWidth - ucBar.Width)
        End If
      Else
        '-- Hide scroll bar
        ucBar.Visible = False
        Call ucBar.Refresh
        Call pvUpdateRectRight(ScaleWidth)
    End If

    '-- Update max value
    ucBar.LargeChange = m_VisibleRows
    ucBar.Max = (UBound(m_List) - m_VisibleRows) + -(Not m_PerfectRowPad)

    On Error GoTo 0
End Sub

Private Sub pvCalculateRects()
  
  Dim nRows    As Integer
  Dim nRct     As Integer
  Dim nItmH    As Integer
  Dim nTxtH    As Integer
  Dim nBarLeft As Integer
  
    nItmH = m_ThumbnailHeight + mdef_ItemHeightPad
    nTxtH = TextHeight(vbNullString)
    nBarLeft = IIf(ucBar.Visible, ucBar.Left, ScaleWidth)
    
    '-- Main rect.
    Call SetRect(m_ControlRect, 0, 0, ScaleWidth, ScaleHeight)
    
    '-- Item rects.
    nRows = m_VisibleRows - 1
    ReDim m_RectExt(nRows)
    ReDim m_RectTxt(nRows)
    ReDim m_RectSel(nRows)
    ReDim m_RectImg(nRows)
    
    For nRct = 0 To nRows
        Call SetRect(m_RectExt(nRct), 0, nRct * nItmH, nBarLeft, nRct * nItmH + nItmH)
        Call SetRect(m_RectTxt(nRct), 1, nRct * nItmH + 1, nBarLeft - 1, nRct * nItmH + nTxtH + 2)
        Call SetRect(m_RectSel(nRct), 1, nRct * nItmH + nTxtH + 2, nBarLeft - 1, nRct * nItmH + nItmH - 1)
        Call SetRect(m_RectImg(nRct), 3, nRct * nItmH + nTxtH + 4, nBarLeft - 3, nRct * nItmH + nItmH - 3)
    Next nRct
    
    '-- Pad rect.
    If (Not m_PerfectRowPad) Then
        With m_RectExt(nRows)
            Call SetRect(m_RectRowPad, 0, .y1, .x2, ScaleHeight)
        End With
      Else
        Call SetRect(m_RectRowPad, 0, 0, 0, 0)
    End If
End Sub

Private Sub pvUpdateRectRight(ByVal New_Right As Integer)

  Dim nRct As Integer
    
    '-- Rects. right offset
    For nRct = 0 To m_VisibleRows - 1
        m_RectExt(nRct).x2 = New_Right
        m_RectTxt(nRct).x2 = New_Right - 1
        m_RectSel(nRct).x2 = New_Right - 1
        m_RectImg(nRct).x2 = New_Right - 3
    Next nRct
    m_RectRowPad.x2 = New_Right
End Sub

Private Sub pvSetColors()
    '-- Get long colors
    Call OleTranslateColor(vbButtonFace, 0, m_ClrButtonFace)
    Call OleTranslateColor(vb3DDKShadow, 0, m_Clr3DDKShadow)
End Sub
