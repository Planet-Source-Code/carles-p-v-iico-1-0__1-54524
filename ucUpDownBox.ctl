VERSION 5.00
Begin VB.UserControl ucUpDownBox 
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1845
   LockControls    =   -1  'True
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   123
   Begin VB.VScrollBar ucBar 
      Height          =   300
      Left            =   1125
      Max             =   -1
      Min             =   1
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   240
   End
   Begin VB.TextBox txtValue 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   0
      Width           =   1125
   End
End
Attribute VB_Name = "ucUpDownBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucUpDownBox.ctl
' Author:        Carles P.V.
' Dependencies:  None
' Last revision: 2003.12.08
'================================================

Option Explicit

'-- API:

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXVSCROLL As Long = &H2

'//

'-- Public Enums.:
Public Enum ebBorderStyleConstants
    [None] = 0
    [3D]
End Enum

'-- Private Variables:
Private m_Min As Long
Private m_Max As Long

'-- Event Declarations:
Public Event Change()



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()
    
    '-- Get system default
    ucBar.Width = GetSystemMetrics(SM_CXVSCROLL)
End Sub

Private Sub UserControl_Resize()

    '-- Adjust width [?]
    If (ScaleWidth < 2 * ucBar.Width) Then
        Width = (2 * ucBar.Width + (Width \ Screen.TwipsPerPixelX - ScaleWidth)) * Screen.TwipsPerPixelX
    End If
    '-- Adjust height
    Height = ((TextHeight("") + 4) + (Height \ Screen.TwipsPerPixelY - ScaleHeight)) * Screen.TwipsPerPixelY
    
    '-- Locate controls
    Call txtValue.Move(1, 2, ScaleWidth - ucBar.Width - 1, ScaleHeight - 2)
    Call ucBar.Move(txtValue.Width + 1, 0, ucBar.Width, ScaleHeight)
End Sub

'========================================================================================
' Text box
'========================================================================================

Private Sub txtValue_GotFocus()

    '-- Select Text box contents
    txtValue.SelStart = 0
    txtValue.SelLength = Len(txtValue)
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '-- Key support
    Select Case KeyCode
        Case vbKeyUp:   KeyCode = 0: ucBar.Value = 1
        Case vbKeyDown: KeyCode = 0: ucBar.Value = -1
    End Select
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    
    '-- Only numbers (allow [KeyBack] and [-])
    If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtValue_Change()
    
    '-- Check
    If (IsNumeric(txtValue)) Then
        
        Select Case True
            Case txtValue < m_Min
                txtValue = m_Min
                txtValue.SelStart = 0
                txtValue.SelLength = Len(txtValue)
            Case txtValue > m_Max
                txtValue = m_Max
                txtValue.SelStart = Len(txtValue)
            Case Else
        End Select
        RaiseEvent Change
      
      Else
        '-- Reset
        txtValue = m_Min
        txtValue.SelStart = 0
        txtValue.SelLength = Len(txtValue)
    End If
End Sub

'========================================================================================
' Scroll bar
'========================================================================================

Private Sub ucBar_Change()

    If (ucBar.Value <> 0) Then
    
        '-- Apply inc.
        Select Case ucBar.Value
            Case Is > 0 '[+1]
                If (txtValue < m_Max) Then txtValue = txtValue + 1
                txtValue.SelStart = 0
                txtValue.SelLength = Len(txtValue)
            Case Is < 0 '[-1]
                If (txtValue > m_Min) Then txtValue = txtValue - 1
                txtValue.SelStart = 0
                txtValue.SelLength = Len(txtValue)
        End Select
        
        '--  Reset Scroll bar
        ucBar.Value = 0
    End If
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get Alignment() As AlignmentConstants
    Alignment = txtValue.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    txtValue.Alignment = New_Alignment
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    txtValue.BackColor = New_BackColor
End Property

Public Property Get BorderStyle() As ebBorderStyleConstants
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As ebBorderStyleConstants)
    UserControl.BorderStyle = New_BorderStyle
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtValue.Enabled = New_Enabled
    ucBar.Enabled = New_Enabled
    UserControl.Enabled = New_Enabled
End Property

Public Property Get Font() As Font
    Set Font = txtValue.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtValue.Font = New_Font
    Set UserControl.Font = New_Font
    UserControl_Resize
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = txtValue.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtValue.ForeColor = New_ForeColor
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property
Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
End Property

Public Property Get Min() As Long
    Min = m_Min
End Property
Public Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
End Property

Public Property Get Value() As Long
Attribute Value.VB_MemberFlags = "400"
    Value = txtValue
End Property
Public Property Let Value(ByVal New_Value As Long)
    If (New_Value < m_Min) Then New_Value = m_Min
    If (New_Value > m_Max) Then New_Value = m_Max
    txtValue = New_Value
End Property

'//

Private Sub UserControl_InitProperties()
    UserControl.BorderStyle = [3D]
    UserControl.BackColor = vbWindowBackground
    Set Font = Ambient.Font
    m_Min = 0
    m_Max = 100
    txtValue = m_Min
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UserControl.BackColor = .ReadProperty("BackColor", vbWindowBackground)
        UserControl.BorderStyle = .ReadProperty("BorderStyle", [3D])
        Enabled = .ReadProperty("Enabled", True)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        txtValue.Alignment = .ReadProperty("Alignment", vbLeftJustify)
        txtValue.BackColor = .ReadProperty("BackColor", vbWindowBackground)
        txtValue.ForeColor = .ReadProperty("ForeColor", vbWindowText)
        Set txtValue.Font = .ReadProperty("Font", Ambient.Font)
        m_Min = .ReadProperty("Min", 0)
        m_Max = .ReadProperty("Max", 100)
        txtValue = m_Min
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Alignment", txtValue.Alignment, vbLeftJustify
        .WriteProperty "BackColor", txtValue.BackColor, vbWindowBackground
        .WriteProperty "BorderStyle", UserControl.BorderStyle, 1
        .WriteProperty "ForeColor", txtValue.ForeColor, vbWindowText
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "Font", txtValue.Font, Ambient.Font
        .WriteProperty "Min", m_Min, 0
        .WriteProperty "Max", m_Max, 100
    End With
End Sub
