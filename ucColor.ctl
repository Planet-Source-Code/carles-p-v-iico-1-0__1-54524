VERSION 5.00
Begin VB.UserControl ucColor 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   LockControls    =   -1  'True
   ScaleHeight     =   54
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   67
   Begin VB.Image iScreen 
      Height          =   150
      Left            =   0
      ToolTipText     =   "Color Screen"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "ucColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucColor.ctl
' Author:        Carles P.V.
' Dependencies:  cDIB.cls
' Last revision: 2004.06.28
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
Private Const BDR_RAISED      As Long = &H5
Private Const BF_RECT         As Long = &HF

Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT2, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDst As Any, ByVal Length As Long, ByVal Fill As Byte)

'//

'-- Default Property Values:
Private Const m_def_IsColorScreen As Boolean = False

'-- Property Variables:
Private m_IsColorScreen   As Boolean
Private m_PickColorCursor As Boolean
Private m_Alpha As Byte
Private m_Color As Long

'-- Private Variables:
Private m_Rect            As RECT2
Private m_ImageRect       As RECT2
Private m_oPatternDIB     As cDIB

'-- Event Declarations:
Public Event Click(ByVal Button As Integer, ByVal Shift As Integer)
Public Event ColorChange()


'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()
    
    '-- Load image / pointer
    iScreen.Picture = LoadResPicture("ICON_COLOR_SCREEN", vbResIcon)
    UserControl.MouseIcon = LoadResPicture("CURSOR_COLORSELECTOR", vbResCursor)

    '-- Initialize pattern DIB
    Set m_oPatternDIB = New cDIB
    
    '-- Opaque
    m_Alpha = 255
End Sub

'//

Private Sub UserControl_Resize()
    
  Dim lScan    As Long
  Dim lPtrScan As Long
  Dim aPattern As Byte
  
    '-- 'Scanline' pattern
    aPattern = &HF0
  
    '-- Build pattern
    If (ScaleWidth > 2 And ScaleHeight > 2) Then
        With m_oPatternDIB
            If (.Create(ScaleWidth - 2, ScaleHeight - 2, [01_bpp])) Then
                For lScan = 0 To ScaleHeight - 3
                    If (lScan Mod 4 = 0) Then
                        aPattern = Not aPattern
                    End If
                    lPtrScan = .lpBits + .BytesPerScanline * lScan
                    Call FillMemory(ByVal lPtrScan, .BytesPerScanline, aPattern)
                Next lScan
            End If
        End With
    End If
    
    '-- Edge rect
    With m_Rect
        .x2 = ScaleWidth
        .y2 = ScaleHeight
    End With
    
    '-- Image rect
    With m_ImageRect
        .x2 = iScreen.Width
        .y2 = iScreen.Height
    End With
    Call iScreen.Move((ScaleWidth - 32) \ 2, (ScaleHeight - 32) \ 2)
    
    '-- Refresh control
    Call Me.Refresh
End Sub

Public Sub Refresh()
    With UserControl
        If (m_oPatternDIB.hDIB) Then
            Call m_oPatternDIB.Stretch(.hDC, 1, 1, ScaleWidth - 2, ScaleHeight - 2)
        End If
        Call DrawEdge(.hDC, m_Rect, BDR_SUNKENOUTER, BF_RECT)
        Call .Refresh
    End With
End Sub

'//

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (PtInRect(m_Rect, x, y) <> 0) Then
        RaiseEvent Click(Button, Shift)
    End If
End Sub

Private Sub iScreen_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (PtInRect(m_ImageRect, x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY) <> 0) Then
        RaiseEvent Click(Button, Shift)
    End If
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
Attribute Color.VB_UserMemId = 0
    Color = m_Color
End Property
Public Property Let Color(ByVal New_BackColor As OLE_COLOR)
    If (New_BackColor > -1) Then
        m_Color = New_BackColor
        Call pvSetPatternColors
        Call Me.Refresh
        RaiseEvent ColorChange
    End If
End Property

Public Property Let Alpha(ByVal New_Alpha As Byte)
    m_Alpha = New_Alpha
    Call pvSetPatternColors
End Property
Public Property Get Alpha() As Byte
    Alpha = m_Alpha
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property

Public Property Get PickColorCursor() As Boolean
    PickColorCursor = m_PickColorCursor
End Property
Public Property Let PickColorCursor(ByVal New_PickColorCursor As Boolean)
    m_PickColorCursor = New_PickColorCursor
    MousePointer = IIf(New_PickColorCursor, vbCustom, vbDefault)
End Property

Public Property Get IsColorScreen() As Boolean
    IsColorScreen = m_IsColorScreen
End Property
Public Property Let IsColorScreen(ByVal New_IsColorScreen As Boolean)
    m_IsColorScreen = New_IsColorScreen
    iScreen.Visible = m_IsColorScreen
End Property

'//

Private Sub UserControl_InitProperties()
    m_IsColorScreen = m_def_IsColorScreen
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Color = .ReadProperty("Color", &H0)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        PickColorCursor = .ReadProperty("PickColorCursor", False)
        IsColorScreen = PropBag.ReadProperty("IsColorScreen", m_def_IsColorScreen)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Color", m_Color, &H0
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "PickColorCursor", m_PickColorCursor, False
        .WriteProperty "IsColorScreen", m_IsColorScreen, m_def_IsColorScreen
    End With
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvSetPatternColors()
  
  Dim R       As Long
  Dim G       As Long
  Dim B       As Long
  Dim aPal(7) As Byte
    
    If (m_oPatternDIB.hDIB) Then
    
        R = (m_Color And &HFF&)
        G = (m_Color And &HFF00&) \ 256
        B = (m_Color And &HFF0000) \ 65536
        
        aPal(0) = (&HC0& * (255 - m_Alpha)) \ 255 + (B * m_Alpha) \ 255
        aPal(1) = (&HC0& * (255 - m_Alpha)) \ 255 + (G * m_Alpha) \ 255
        aPal(2) = (&HC0& * (255 - m_Alpha)) \ 255 + (R * m_Alpha) \ 255
        
        aPal(4) = (&HFF& * (255 - m_Alpha)) \ 255 + (B * m_Alpha) \ 255
        aPal(5) = (&HFF& * (255 - m_Alpha)) \ 255 + (G * m_Alpha) \ 255
        aPal(6) = (&HFF& * (255 - m_Alpha)) \ 255 + (R * m_Alpha) \ 255
    
        Call m_oPatternDIB.SetPalette(aPal())
    End If
End Sub
