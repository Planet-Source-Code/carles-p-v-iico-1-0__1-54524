Attribute VB_Name = "mMain"
Option Explicit

'-- API :

Private Const GWL_WNDPROC          As Long = (-4)
Private Const WM_WINDOWPOSCHANGING As Long = &H46
Private Const WM_CTLCOLORSCROLLBAR As Long = &H137
Private Const WM_MOUSEWHEEL        As Long = &H20A
Private Const WM_CANCELMODE        As Long = &H1F

Private Const WM_DRAWCLIPBOARD     As Long = &H308
Private Const WM_CHANGECBCHAIN     As Long = &H30D

Private Type WINDOWPOS
    hWnd       As Long
    hWndInsAft As Long
    x          As Long
    y          As Long
    cX         As Long
    cY         As Long
    Flags      As Long
End Type

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ChangeClipboardChain Lib "user32" (ByVal hWnd As Long, ByVal hWndNext As Long) As Long

Private Const LF_FACESIZE       As Long = 32
Private Const LF_FULLFACESIZE   As Long = 64
Private Const TRUETYPE_FONTTYPE As Long = 4

Private Type LOGFONT
    lfHeight                As Long
    lfWidth                 As Long
    lfEscapement            As Long
    lfOrientation           As Long
    lfWeight                As Long
    lfItalic                As Byte
    lfUnderline             As Byte
    lfStrikeOut             As Byte
    lfCharSet               As Byte
    lfOutPrecision          As Byte
    lfClipPrecision         As Byte
    lfQuality               As Byte
    lfPitchAndFamily        As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type ENUMLOGFONTEX
    elfLogFont                   As LOGFONT
    elfFullName(LF_FULLFACESIZE) As Byte
    elfStyle(LF_FACESIZE)        As Byte
    elfScript(LF_FACESIZE)       As Byte
End Type

Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long

'//

'-- Private Constants:
Private Const MIN_WIDTH  As Long = 557
Private Const MAX_WIDTH  As Long = 10000
Private Const MIN_HEIGHT As Long = 451
Private Const MAX_HEIGHT As Long = 10000
Private Const TOP_MRG    As Long = 35
Private Const TAB_MRG    As Long = 5
  
'-- Private Variables:
Private m_lOldfMainProc      As Long
Private m_lOlducIconListProc As Long
Private m_ClipboardChain     As Long

'-- Private Types:
Private Type COLORINFO
    Palette() As Byte
    ColorIdxA As Integer
    ColorIdxB As Integer
End Type

'-- Global Variables:
Public G_GDIpToken    As Long
Public G_OpenSavePath As String
Public G_ImportPath   As String
Public G_PalettePath  As String
Public G_oICON        As cIcon
Public G_ImageIdx     As Integer
Public G_ColorInfo()  As COLORINFO
Public G_ColorScreen  As Long
Public G_LastTool     As icToolKeyCts

'//

Public Sub InitializeApp()

  Dim GpInput As GdiplusStartupInput
  Dim nIdx    As Integer
    
    On Error GoTo errH
    
    '-- Load the GDI+ library
    GpInput.GdiplusVersion = 1
    If (mGDIp.GdiplusStartup(G_GDIpToken, GpInput) <> [OK]) Then
        GoTo errH
    End If
    
    On Error GoTo 0

    With fMain
         
         '-- App. icon
        .Icon = LoadResPicture("ICO", vbResIcon)
        '-- Icon object
        Set G_oICON = New cIcon
        '-- Undo module
        Call mUndo.InitializeIDs
                
        '-- Toolbars
        With .ucToolbarMain
            Call .BuildToolbar(LoadResPicture("TB_MAIN", vbResBitmap), &HFF00FF, 16, 16, "BBB|BB|BBB|B")
            Call .SetTooltips("New|Load|Save as...|Undo|Redo|Cut|Copy|Paste|Screen capture")
        End With
        With .ucToolbarFormat
            Call .BuildToolbar(LoadResPicture("TB_FORMAT", vbResBitmap), &HFF00FF, 18, 18, "BB")
            Call .SetTooltips("Add format|Remove format")
        End With
        With .ucToolbarDrawTools
            Call .BuildToolbar(LoadResPicture("TB_DRAWTOOLS", vbResBitmap), &HFF00FF, 18, 18, "OOOOOOOOO")
            Call .SetTooltips("Selection frame|Pencil|Straight line|Brush|Shape|Flood fill|Text|Color selector|Set hot spot")
            Call .CheckButton(2, True)
        End With
        With .ucToolbarPixelMask
            Call .BuildToolbar(LoadResPicture("TB_PIXELMASK", vbResBitmap), &HFF00FF, 11, 13, "OOO")
            Call .SetTooltips("All|Even|Odd")
        End With
        With .ucToolbarFont
            Call .BuildToolbar(LoadResPicture("TB_FONT", vbResBitmap), &HFF00FF, 11, 13, "CC")
            Call .SetTooltips("Bold|Italic")
        End With
        
        '-- Icon image/s list
        Call .ucIconList.SetIconSource(G_oICON)
        .ucIconList.ColorScreen = G_ColorScreen
        
        '-- Icon canvas
        Call .ucIconCanvas.SetIconSource(G_oICON)
        .ucIconCanvas.CanvasEx.ColorScreen = G_ColorScreen
        .ucIconCanvas.Tool = [icPencil]: G_LastTool = [icPencil]
        
        '-- Preview window
        .picPreview.BackColor = G_ColorScreen
        
        '-- Screen color picker
        .ucColorScreen.Color = G_ColorScreen
        
        '-- Alpha
        .ucAlphaPicker.Max = 255
        .ucAlphaPicker.Value = 255
        
        '== Draw tools
        
        '-- Pixels
        .fraDrawTools(2).Visible = True
        Call .ucToolbarPixelMask.CheckButton(1, True)
        
        '-- Straight line / brush
        For nIdx = 1 To 9 Step 2
            Call .cbStraightLineWidth.AddItem(nIdx)
            Call .cbBrushLineWidth.AddItem(nIdx)
            Call .cbShapeLineWidth.AddItem(nIdx)
        Next nIdx
        .cbStraightLineWidth.ListIndex = 0
        .cbBrushLineWidth.ListIndex = 0
        .cbShapeLineWidth.ListIndex = 0
        
        '-- Shape
        Call .cbShape.AddItem("Rectangle")
        Call .cbShape.AddItem("Rectangle solid")
        Call .cbShape.AddItem("Ellipse")
        Call .cbShape.AddItem("Ellipse solid")
        .cbShape.ListIndex = 0
        
        '-- Text
        '-  Font sizes
        For nIdx = 5 To 72
            Call .cbFontSize.AddItem(nIdx)
        Next nIdx
        .cbFontSize.ListIndex = 3
        '-  Font names
        Call EnumFontFamilies(.hDC, vbNullString, AddressOf pvEnumFontFamProc, ByVal 0)
        For nIdx = 0 To .cbFont.ListCount - 1
            If (.cbFont.List(nIdx) = "Tahoma") Then
                .cbFont.ListIndex = nIdx
                Exit For
            End If
        Next nIdx
        Call mMisc.ChangeDropDownSize(.cbFont, 200, 300)
        
        '-- New style
        Call mMisc.ChangeBorderStyle(.ucIconList.hWnd, [bsThin])
        Call mMisc.ChangeBorderStyle(.ucIconCanvas.hWnd, [bsThin])
        Call mMisc.ChangeBorderStyle(.picPreview.hWnd, [bsThin])
        
        '-- Subclassed...
        DoEvents
        Call pvSubclass_fMain
        Call pvSubclass_ucIconList
    End With
    Exit Sub

errH:
    Call MsgBox("Error loading GDI+!", vbCritical)
    Call TerminateApp
    On Error GoTo 0
End Sub

Public Sub TerminateApp()
    
    '-- Unload the GDI+ Dll
    If (G_GDIpToken <> 0) Then
        Call mGDIp.GdiplusShutdown(G_GDIpToken)
    End If
    
    '-- Unsubclass
    If (m_ClipboardChain <> 0) Then
        Call ChangeClipboardChain(fMain.hWnd, m_ClipboardChain)
    End If
    If (m_lOldfMainProc <> 0) Then
        Call SetWindowLong(fMain.hWnd, GWL_WNDPROC, m_lOldfMainProc)
    End If
    If (m_lOlducIconListProc <> 0) Then
        Call SetWindowLong(fMain.ucIconList.hWnd, GWL_WNDPROC, m_lOlducIconListProc)
    End If
    
    '-- Unload/free forms
    Call Unload(fCapture)
    Set fCapture = Nothing
    Call Unload(fImport)
    Set fImport = Nothing
    Call Unload(fNew)
    Set fNew = Nothing
    Call Unload(fNewFormat)
    Set fNewFormat = Nothing
    Call Unload(fMaskFix)
    Set fMaskFix = Nothing
    Set fMain = Nothing
End Sub

'//

Public Sub fMain_Resize()
    
  Dim nFrm As Integer
  
    On Error Resume Next
    
    With fMain
        
        If (.WindowState <> vbMinimized) Then
            
            '-- Locate, resize controls
            Call .ucToolbarFormat.Move(TAB_MRG - 1, TOP_MRG - 1)
            Call .ucIconList.Move(TAB_MRG, TOP_MRG + .ucToolbarFormat.Height + TAB_MRG, .ucIconList.Width, .ScaleHeight - (TOP_MRG + .ucToolbarFormat.Height + .picPreview.Height + .ucIconInfo.Height + 3 * TAB_MRG))
            Call .picPreview.Move(TAB_MRG, .ucIconList.Top + .ucIconList.Height + TAB_MRG)
            Call .ucColorScreen.Move(.ucIconList.Left + .ucIconList.Width + 2 * TAB_MRG, TOP_MRG)
            Call .ucColorA.Move(.ucColorScreen.Left + .ucColorScreen.Width + 2 * TAB_MRG, TOP_MRG)
            Call .ucColorB.Move(.ucColorA.Left + .ucColorA.Width + TAB_MRG \ 2, TOP_MRG)
            Call .ucAlphaPicker.Move(.ucColorScreen.Left, .ucIconList.Top, .ucPalettePicker.Width)
            Call .ucPalettePicker.Move(.ucColorScreen.Left, .ucAlphaPicker.Top + .ucAlphaPicker.Height + TAB_MRG, 0, .ScaleHeight - (.ucAlphaPicker.Top + .ucAlphaPicker.Height + .ucIconInfo.Height + 2 * TAB_MRG))
            Call .ucToolbarDrawTools.Move(.ucColorB.Left + .ucColorB.Width + 2 * TAB_MRG - 1, .ucColorB.Top - 1)
            Call .ucIconCanvas.Move(.ucToolbarDrawTools.Left + 1, .ucPalettePicker.Top, .ScaleWidth - (.ucPalettePicker.Left + .ucPalettePicker.Width + 3 * TAB_MRG), .ScaleHeight - (.ucAlphaPicker.Top + .ucAlphaPicker.Height + .ucIconInfo.Height + 2 * TAB_MRG))
            
            For nFrm = 1 To .fraDrawTools.Count
                Call .fraDrawTools(nFrm).Move(.ucToolbarDrawTools.Left + 1, .ucIconList.Top, .fraDrawTools(nFrm).Width, .ucAlphaPicker.Height)
            Next nFrm
        End If
    End With
    
    On Error GoTo 0
End Sub

'========================================================================================
' Private (Subclassing procedures)
'========================================================================================

Public Sub pvSubclass_fMain()
    
    '-- New Main proc.
    m_lOldfMainProc = SetWindowLong(fMain.hWnd, GWL_WNDPROC, AddressOf pvfMainProc)
    m_ClipboardChain = SetClipboardViewer(fMain.hWnd)
End Sub

Public Sub pvSubclass_ucIconList()
    
    '-- New Icon List proc.
    m_lOlducIconListProc = SetWindowLong(fMain.ucIconList.hWnd, GWL_WNDPROC, AddressOf pvucIconListProc)
End Sub


Private Function pvfMainProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim uWP As WINDOWPOS

    Select Case Msg
    
        Case WM_DRAWCLIPBOARD
            
            '-- Clipboard has been changed
            If (wParam = m_ClipboardChain) Then
                m_ClipboardChain = lParam
            ElseIf (m_ClipboardChain <> 0) Then
                Call SendMessage(m_ClipboardChain, Msg, wParam, lParam)
            End If
            
            With fMain
                If (Clipboard.GetFormat(vbCFDIB)) Then
                    Call .ucIconCanvas.DestroyPrivateClipboard
                    Call .ucToolbarMain.EnableButton(8, True)
                  Else
                    Call .ucToolbarMain.EnableButton(8, False)
                End If
                Call .frUpdateAlphaTools
            End With
       
        Case WM_CHANGECBCHAIN
        
            '-- Some other viewer is being removed from chain
            If (m_ClipboardChain <> 0) Then
                Call SendMessage(m_ClipboardChain, Msg, wParam, lParam)
            End If
           
        Case WM_WINDOWPOSCHANGING
            
            '-- Window position is changing
            Call CopyMemory(uWP, ByVal lParam, Len(uWP))
            With uWP
                If (.cX < MIN_WIDTH) Then .cX = MIN_WIDTH
                If (.cX > MAX_WIDTH) Then .cX = MAX_WIDTH
                If (.cY < MIN_HEIGHT) Then .cY = MIN_HEIGHT
                If (.cY > MAX_HEIGHT) Then .cY = MAX_HEIGHT
            End With
            Call CopyMemory(ByVal lParam, uWP, Len(uWP))
        
        Case WM_CTLCOLORSCROLLBAR
            
            '-- Skip
            pvfMainProc = 0
            Exit Function
    End Select
    
    pvfMainProc = CallWindowProc(m_lOldfMainProc, hWnd, Msg, wParam, lParam)
End Function

Private Function pvucIconListProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case Msg
    
        Case WM_CTLCOLORSCROLLBAR
        
            '-- Skip
            Exit Function
        
        Case WM_MOUSEWHEEL
        
            '-- Scroll list
            With fMain.ucIconList
                If (.ListCount) Then
                    Select Case wParam
                        Case Is > 0
                            If (.ListIndex > 0) Then
                                .ListIndex = .ListIndex - 1
                            End If
                        Case Else
                            If (.ListIndex < .ListCount - 1) Then
                                .ListIndex = .ListIndex + 1
                            End If
                    End Select
                End If
            End With
    End Select
    
    pvucIconListProc = CallWindowProc(m_lOlducIconListProc, hWnd, Msg, wParam, lParam)
End Function

'========================================================================================
' Private (Font enum.)
'========================================================================================

Private Function pvEnumFontFamProc(ByVal lpFG As Long, ByVal lpNTM As Long, ByVal iFontType As Long, lParam As Long) As Long
   
  Dim uLFEx     As ENUMLOGFONTEX
  Dim sFacename As String
  Dim lPos      As Long
  
    If (iFontType And TRUETYPE_FONTTYPE) Then
    
        Call CopyMemory(uLFEx, ByVal lpFG, LenB(uLFEx))
        
        sFacename = StrConv(uLFEx.elfLogFont.lfFaceName, vbUnicode)
        lPos = InStr(sFacename, Chr$(0))
        If (lPos > 0) Then
            sFacename = Left$(sFacename, (lPos - 1))
        End If
        Call fMain.cbFont.AddItem(sFacename)
    End If
    
    pvEnumFontFamProc = 1
End Function
