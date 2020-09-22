VERSION 5.00
Begin VB.UserControl ucSlider 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   ClipControls    =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   19
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   169
End
Attribute VB_Name = "ucSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucSlider.ctl (reduced)
' Author:        Carles P.V.
' Dependencies:  None
' Last revision: 2003.12.08
'================================================

Option Explicit

'-- API:

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type TRIVERTEX
    x     As Long
    y     As Long
    Red   As Integer
    Green As Integer
    Blue  As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft  As Long
    LowerRight As Long
End Type

Private Const COLOR_BTNFACE        As Long = 15
Private Const BDR_SUNKENOUTER      As Long = &H2
Private Const BDR_RAISED           As Long = &H5
Private Const BF_RECT              As Long = &HF
Private Const GRADIENT_FILL_RECT_H As Long = 0
Private Const GRADIENT_FILL_RECT_V As Long = 1

Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT2, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long

'-- Private Constants:
Private Const THUMB_WIDTH As Long = 8

'-- Private Variables:
Private m_rctControl  As RECT2
Private m_rctThumb    As RECT2
Private m_ThumbOffset As Long
Private m_bInThumb    As Boolean
Private m_LastValue   As Byte

'-- Property Variables:
Private m_Value As Long
Private m_Max   As Long

'-- Event Declarations:
Public Event Change(ByVal Value As Byte)
Public Event MouseDown()
Public Event MouseUp()



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()
    
    '-- Max default
    m_Max = 100
    '-- Mouse pointer
    Set UserControl.MouseIcon = LoadResPicture("CURSOR_HAND", vbResCursor)
End Sub

Private Sub UserControl_Resize()
    
    Call SetRect(m_rctControl, 0, 0, ScaleWidth, ScaleHeight)
    With m_rctThumb
        .x1 = m_Value * ((ScaleWidth - 1) - THUMB_WIDTH) / 255 + 1
        .x2 = .x1 + THUMB_WIDTH - 1
        .y1 = 1
        .y2 = ScaleHeight - 1
    End With
    Call Me.Refresh
End Sub

Public Sub Refresh()
    
    With UserControl
        If (Me.Enabled) Then
            Call pvFillGradient(m_rctControl, vbWhite, vbBlack, False)
            Call FillRect(.hDC, m_rctThumb, GetSysColorBrush(COLOR_BTNFACE))
            Call DrawEdge(.hDC, m_rctThumb, BDR_RAISED, BF_RECT)
          Else
            Call FillRect(.hDC, m_rctControl, GetSysColorBrush(COLOR_BTNFACE))
        End If
        Call DrawEdge(.hDC, m_rctControl, BDR_SUNKENOUTER, BF_RECT)
        Call .Refresh
    End With
End Sub

'========================================================================================
' Thumb control
'========================================================================================

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
  Dim sRet As String
  
    If (Me.Enabled) Then
    
        Select Case Button
        
            Case vbLeftButton
        
                With m_rctThumb
                    
                    If (Button = vbLeftButton) Then
                       
                        m_bInThumb = True
                        
                        If (PtInRect(m_rctThumb, x, y)) Then
                            m_ThumbOffset = x - .x1
                          Else
                            m_ThumbOffset = THUMB_WIDTH / 2
                            UserControl_MouseMove Button, Shift, x, y
                        End If
                    End If
                    RaiseEvent MouseDown
                End With
                
            Case vbRightButton
            
                sRet = InputBox("Enter new value [0-" & m_Max & "]", , m_Value)
                If (IsNumeric(sRet)) Then
                    If (Val(sRet) >= 0 And Val(sRet) <= m_Max) Then
                        Value = Val(sRet)
                    End If
                End If
        End Select
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (m_bInThumb) Then
        
        With m_rctThumb
            
            If (x - m_ThumbOffset < 0) Then
                .x1 = 1
              ElseIf (x - m_ThumbOffset > (ScaleWidth - 1) - THUMB_WIDTH) Then
                .x1 = (ScaleWidth - 1) - THUMB_WIDTH + 1
              Else
                .x1 = x - m_ThumbOffset + 1
            End If
            .x2 = .x1 + THUMB_WIDTH - 1
        
            Let Value = ((.x1 - 1) / ((ScaleWidth - 1) - THUMB_WIDTH) * 255)
        End With
        
        Call Me.Refresh
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (Button = vbLeftButton) Then RaiseEvent MouseUp
    m_bInThumb = False
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Let Value(ByVal New_Value As Byte)
    
    m_Value = New_Value
        
    If (m_Value <> m_LastValue) Then
        m_LastValue = New_Value
        
        If (Not m_bInThumb) Then
            With m_rctThumb
                .x1 = m_Value * ((ScaleWidth - 1) - THUMB_WIDTH) / m_Max + 1
                .x2 = .x1 + THUMB_WIDTH - 1
            End With
        End If
        
        Call Me.Refresh
        RaiseEvent Change(New_Value)
    End If
End Property
Public Property Get Value() As Byte
    Value = m_Value
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
End Property
Public Property Get Max() As Long
    Max = m_Max
    Let Value = m_Value
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled: Call Me.Refresh
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

'========================================================================================
' Properties
'========================================================================================

Private Sub pvFillGradient(lpRect As RECT2, _
                           ByVal clrFirst As OLE_COLOR, _
                           ByVal clrSecond As OLE_COLOR, _
                           Optional ByVal bVertical As Boolean)
            
  Dim pVert(0 To 1) As TRIVERTEX
  Dim clr           As OLE_COLOR
  Dim pGradRect     As GRADIENT_RECT
    
    Call OleTranslateColor(clrFirst, 0, clr)
    With pVert(0)
        .x = lpRect.x1
        .y = lpRect.y1
        .Red = pvRed(clr)
        .Green = pvGreen(clr)
        .Blue = pvBlue(clr)
    End With
    Call OleTranslateColor(clrSecond, 0, clr)
    With pVert(1)
        .x = lpRect.x2
        .y = lpRect.y2
        .Red = pvRed(clr)
        .Green = pvGreen(clr)
        .Blue = pvBlue(clr)
    End With
    With pGradRect
        .UpperLeft = 0
        .LowerRight = 1
    End With
    Call GradientFill(UserControl.hDC, pVert(0), 2, pGradRect, 1, IIf(Not bVertical, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V))
End Sub

Private Function pvRed(ByVal clr As OLE_COLOR) As Long
    pvRed = ((clr \ &H1) And &HFF) * &H100&
    If (pvRed >= &H8000&) Then
        pvRed = pvRed - &H10000
    End If
End Function
    
Private Function pvGreen(ByVal clr As OLE_COLOR) As Long
    pvGreen = ((clr \ &H100) And &HFF) * &H100&
    If (pvGreen >= &H8000&) Then
        pvGreen = pvGreen - &H10000
    End If
End Function

Private Function pvBlue(ByVal clr As OLE_COLOR) As Long
    pvBlue = ((clr \ &H10000) And &HFF) * &H100&
    If (pvBlue >= &H8000&) Then
        pvBlue = pvBlue - &H10000
    End If
End Function

