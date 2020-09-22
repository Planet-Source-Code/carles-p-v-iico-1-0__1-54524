Attribute VB_Name = "mMisc"
Option Explicit

Public Enum eBorderStyleConstants
    [bsNone] = 0
    [bsThin] = 1
    [bsThick] = 2
End Enum

'-- API :

Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDst As Any, ByVal Length As Long, ByVal Fill As Byte)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'//

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_DRAWFRAME     As Long = &H20
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOREDRAW      As Long = &H8
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_SHOWWINDOW    As Long = &H40
Private Const SWP_FLAGS         As Long = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE         As Long = (-16)
Private Const WS_THICKFRAME     As Long = &H40000
Private Const WS_BORDER         As Long = &H800000
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const WS_EX_WINDOWEDGE  As Long = &H100&
Private Const WS_EX_CLIENTEDGE  As Long = &H200&
Private Const WS_EX_STATICEDGE  As Long = &H20000

'//

Private Const CB_SETDROPPEDWIDTH = &H160

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'//

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW As Long = 5

'//

Public Sub RemoveButtonBorderEnhance(oButton As CommandButton)
  
  Dim lRet As Long
  
    lRet = GetWindowLong(oButton.hWnd, GWL_STYLE)
    Call SetWindowLong(oButton.hWnd, GWL_STYLE, lRet And Not &HB)
End Sub

Public Sub ChangeBorderStyle(ByVal hWnd As Long, ByVal eStyle As eBorderStyleConstants)
    
    Select Case eStyle
        Case [bsNone]
            Call pvSetWinStyle(hWnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call pvSetWinStyle(hWnd, GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE)
        Case [bsThin]
            Call pvSetWinStyle(hWnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call pvSetWinStyle(hWnd, GWL_EXSTYLE, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE)
        Case [bsThick]
            Call pvSetWinStyle(hWnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call pvSetWinStyle(hWnd, GWL_EXSTYLE, WS_EX_CLIENTEDGE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE)
    End Select
End Sub

Public Sub ChangeDropDownSize(oCombo As ComboBox, ByVal lWidth As Long, ByVal lHeight As Long)
    
    With oCombo
        '-- Drop down list width
        Call SendMessage(.hWnd, CB_SETDROPPEDWIDTH, lWidth, ByVal 0)
        '-- Drop down list height
        Call MoveWindow(.hWnd, _
                        .Left \ Screen.TwipsPerPixelX, _
                        .Top \ Screen.TwipsPerPixelY, _
                        .Width \ Screen.TwipsPerPixelX, _
                        lHeight, 0)
    End With
End Sub

'//

Public Sub Navigate(ByVal lhWnd As Long, ByVal sURL As String)
    
    '-- Open URL
    Call ShellExecute(lhWnd, "open", sURL, vbNullString, vbNullString, SW_SHOW)
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvSetWinStyle(ByVal lhWnd As Long, ByVal lType As Long, ByVal lStyle As Long, ByVal lStyleNot As Long)

  Dim lRet As Long
    
    lRet = GetWindowLong(lhWnd, lType)
    lRet = (lRet And Not lStyleNot) Or lStyle
    Call SetWindowLong(lhWnd, lType, lRet)
    Call SetWindowPos(lhWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub


