Attribute VB_Name = "mDialogColor"
'================================================
' Module:        mDialogColor.bas
' Author:
' Dependencies:  None
' Last revision: 2004.06.14
'================================================

Option Explicit

Private Type tChooseColor
    lStructSize    As Long
    hwndOwner      As Long
    hInstance      As Long
    rgbResult      As Long
    lpCustColors   As Long
    Flags          As Long
    lCustData      As Long
    lpfnHook       As Long
    lpTemplateName As String
End Type

Private Const CC_RGBINIT   As Long = &H1
Private Const CC_FULLOPEN  As Long = &H2
Private Const CC_ANYCOLOR  As Long = &H100

Private Const CC_NORMAL    As Long = CC_ANYCOLOR Or CC_RGBINIT
Private Const CC_EXTENDED  As Long = CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN

Private Declare Function ChooseColor Lib "comdlg32" Alias "ChooseColorA" (Color As tChooseColor) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long

Private m_CustomColors(15) As Long
Private m_Initialized      As Boolean



Public Function SelectColor(ByVal hWndParent As Long, ByVal DefaultColor As Long, Optional ByVal ShowDlgEx As Boolean = 0) As Long
 
  Dim uCC  As tChooseColor
  Dim lRet As Long
  
  Dim lIdx As Long
  Dim lClr As Long
 
    With uCC
        
        '-- Initiliaze custom colors
        If (m_Initialized = False) Then
            m_Initialized = True
            
            m_CustomColors(0) = &HFAC896
            m_CustomColors(1) = G_ColorScreen
            Call OleTranslateColor(vb3DFace, 0, lClr)
            For lIdx = 2 To 15
                m_CustomColors(lIdx) = lClr
            Next lIdx
        End If
        
        '-- Prepare struct.
        .lStructSize = Len(uCC)
        .hwndOwner = hWndParent
        .rgbResult = DefaultColor
        .lpCustColors = VarPtr(m_CustomColors(0))
        .Flags = IIf(ShowDlgEx, CC_EXTENDED, CC_NORMAL)
        
        '-- Show Color dialog
        lRet = ChooseColor(uCC)
         
        '-- Get color / Cancel
        If (lRet) Then
            SelectColor = .rgbResult
          Else
            SelectColor = -1
        End If
    End With
End Function
