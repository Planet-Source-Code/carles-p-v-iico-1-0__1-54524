Attribute VB_Name = "mUndo"
Option Explicit

'-- Private Constants:
Private Const m_UNDOLEVELS As Long = 25 ' Max Undo levels (per format)

'-- Private Types:
Private Type UNDOINFO
    ImageIdx     As Integer
    UndoPos      As Integer
    UndoMax      As Integer
    Irreversible As Boolean
End Type

Private Type IMAGEHEADER
    Width        As Integer
    Height       As Integer
    BPP          As Byte
    HotSpotX     As Integer
    HotSpotY     As Integer
End Type

'-- Private Variables:
Private m_AppID        As Long     ' Application ID (fMain.hwnd)
Private m_sTemp        As String   ' Temporary folder
Private m_uInfo()      As UNDOINFO ' Undo info
Private m_ListSaved    As Boolean  ' Just saved



'========================================================================================
' Methods
'========================================================================================

Public Sub InitializeIDs()
    
    '-- Application ID and temp. folder
    m_AppID = App.ThreadID
    m_sTemp = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    
    '-- Initialize Undo info array
    ReDim m_uInfo(0)
End Sub

'//

Public Sub InitializeListUndoHistory()
    
  Dim nIdx As Integer
    
    '-- Reset Undo info
    ReDim m_uInfo(G_oICON.Count - 1)
    
    '-- Save first state of each format
    For nIdx = 0 To G_oICON.Count - 1
        Call SaveUndo(nIdx)
    Next nIdx
End Sub

Public Sub CleanListUndoHistory()
    
    '-- Delete all Undo files
    On Error Resume Next
       Call Kill(pvPathList)
    On Error GoTo 0
    
    '-- Reset Undo info
    ReDim m_uInfo(0)
End Sub

'//

Public Sub InsertUndoItemHistory(ByVal ImageIdx As Integer)
    
  Dim nIdx As Integer
  Dim lIdx As Long
    
    '-- Update files (rename)
    For nIdx = UBound(m_uInfo()) To ImageIdx Step -1
        For lIdx = m_uInfo(nIdx).UndoMax - 1 To 0 Step -1
            Name pvPathItem(nIdx, lIdx) As pvPathItem(nIdx + 1, lIdx)
        Next lIdx
    Next nIdx
    
    '-- Update info (move)
    ReDim Preserve m_uInfo(UBound(m_uInfo()) + 1)
    For nIdx = UBound(m_uInfo()) To ImageIdx + 1 Step -1
        m_uInfo(nIdx) = m_uInfo(nIdx - 1)
    Next nIdx
    
    '-- Reset current info item
    m_uInfo(ImageIdx).UndoPos = 0
    m_uInfo(ImageIdx).UndoMax = 0
    m_uInfo(ImageIdx).Irreversible = False
    
    '-- Finaly, save 'first state'
    Call SaveUndo(ImageIdx)
End Sub

Public Sub RemoveUndoItemHistory(ByVal ImageIdx As Integer)
    
  Dim nIdx As Integer
  Dim lIdx As Long
    
    '-- Update files (rename)
    Call Kill(pvPathItem(ImageIdx, -1))
    For nIdx = ImageIdx + 1 To UBound(m_uInfo())
        For lIdx = 0 To m_uInfo(nIdx).UndoMax - 1
            Name pvPathItem(nIdx, lIdx) As pvPathItem(nIdx - 1, lIdx)
        Next lIdx
    Next nIdx
    
    '-- Update info (move)
    For nIdx = ImageIdx To UBound(m_uInfo()) - 1
        m_uInfo(nIdx) = m_uInfo(nIdx + 1)
    Next nIdx
    ReDim Preserve m_uInfo(UBound(m_uInfo()) - 1)
End Sub

'//

Public Sub SaveUndo(ByVal ImageIdx As Integer)

  Dim nIdx As Integer
  
    '-- Save
    Call pvSaveUndoFile(ImageIdx, m_uInfo(ImageIdx).UndoPos)
    m_ListSaved = False
        
    With m_uInfo(ImageIdx)
        
        If (.UndoMax - .UndoPos > 0) Then
            On Error Resume Next
            For nIdx = .UndoPos + 1 To .UndoMax
                Call Kill(pvPathItem(ImageIdx, nIdx))
            Next nIdx
            On Error GoTo 0
        End If
        
        If (.UndoPos < m_UNDOLEVELS) Then
            .UndoPos = .UndoPos + 1
            .UndoMax = .UndoPos
          Else
            Call pvRotateUndoFiles(ImageIdx)
        End If
    End With
End Sub

Public Sub Undo(ByVal ImageIdx As Integer)
    
    With m_uInfo(ImageIdx)
        
        If (.UndoPos > 1) Then
            .UndoPos = .UndoPos - 1
            
            '-- Load
            Call pvLoadUndoFile(ImageIdx, m_uInfo(ImageIdx).UndoPos - 1)
        End If
    End With
End Sub

Public Sub Redo(ByVal ImageIdx As Integer)

    With m_uInfo(ImageIdx)
    
        If (.UndoPos < .UndoMax) Then
            .UndoPos = .UndoPos + 1
            .UndoMax = IIf(.UndoPos > .UndoMax, .UndoPos, .UndoMax)
        
            '-- Load
            Call pvLoadUndoFile(ImageIdx, m_uInfo(ImageIdx).UndoPos - 1)
        End If
    End With
End Sub

'//

Public Function IsItemUndoAvailable(ByVal ImageIdx As Integer) As Boolean
    IsItemUndoAvailable = (m_uInfo(ImageIdx).UndoPos > 1)
End Function

Public Function IsItemRedoAvailable(ByVal ImageIdx As Integer) As Boolean
    IsItemRedoAvailable = (m_uInfo(ImageIdx).UndoPos < m_uInfo(ImageIdx).UndoMax)
End Function

Public Function IsListUndoAvailable() As Boolean

  Dim nIdx As Integer
  
    For nIdx = 0 To UBound(m_uInfo)
        IsListUndoAvailable = (m_uInfo(nIdx).UndoPos > 1)
        If (IsListUndoAvailable) Then Exit For
    Next nIdx
End Function

Public Function IsListIrreversible() As Boolean

  Dim nIdx As Integer
  
    For nIdx = 0 To UBound(m_uInfo)
        IsListIrreversible = m_uInfo(nIdx).Irreversible
        If (IsListIrreversible) Then Exit For
    Next nIdx
End Function

Public Sub SetListSaved(ByVal Saved As Boolean)
    m_ListSaved = Saved
End Sub
Public Function IsListSaved() As Boolean
    IsListSaved = m_ListSaved
End Function

'========================================================================================
' Private
'========================================================================================

Private Sub pvSaveUndoFile(ByVal ImageIdx As Integer, ByVal UndoPos As Integer)

  Dim hFile   As Integer
  Dim uIH     As IMAGEHEADER
  Dim aData() As Byte
    
    '-- Save Undo
    hFile = FreeFile()
    Open pvPathItem(ImageIdx, UndoPos) For Binary Access Write As #hFile
                  
        '-- Write image format header
        With uIH
            .Width = G_oICON.Width(ImageIdx)
            .Height = G_oICON.Height(ImageIdx)
            .BPP = G_oICON.BPP(ImageIdx)
            .HotSpotX = G_oICON.HotSpotX(ImageIdx)
            .HotSpotY = G_oICON.HotSpotY(ImageIdx)
        End With
        Put #hFile, , uIH
        
        '-- Write color info
        Put #hFile, , G_ColorInfo(ImageIdx).Palette()
        Put #hFile, , G_ColorInfo(ImageIdx).ColorIdxA
        Put #hFile, , G_ColorInfo(ImageIdx).ColorIdxB
        
        '-- Write XOR data
        ReDim aData(G_oICON.oXORDIB(ImageIdx).Size - 1)
        Call CopyMemory(aData(0), ByVal G_oICON.oXORDIB(ImageIdx).lpBits, G_oICON.oXORDIB(ImageIdx).Size)
        Put #hFile, , aData()
        
        '-- Write AND data
        ReDim aData(G_oICON.oANDDIB(ImageIdx).Size - 1)
        Call CopyMemory(aData(0), ByVal G_oICON.oANDDIB(ImageIdx).lpBits, G_oICON.oANDDIB(ImageIdx).Size)
        Put #hFile, , aData()
        
    Close #hFile
End Sub

Private Sub pvLoadUndoFile(ByVal ImageIdx As Integer, ByVal UndoPos As Integer)

  Dim hFile   As Integer
  Dim uIH     As IMAGEHEADER
  Dim aData() As Byte
  
    ReDim aPalAND(7) As Byte
    Call FillMemory(aPalAND(4), 3, &HFF)
    
    hFile = FreeFile()
    Open pvPathItem(ImageIdx, UndoPos) For Binary Access Read As #hFile
            
        '-- Read image format header
        Get #hFile, , uIH
        With uIH
        
            '-- Set data format
            Call G_oICON.oXORDIB(ImageIdx).Create(.Width, .Height, .BPP)
            Call G_oICON.oANDDIB(ImageIdx).Create(.Width, .Height, [01_bpp])
            
            '-- Just in case of cursor
            G_oICON.HotSpotX(ImageIdx) = .HotSpotX
            G_oICON.HotSpotY(ImageIdx) = .HotSpotY
            
            '-- Prepare palette array
            Select Case .BPP
                Case Is <= 8
                    ReDim G_ColorInfo(ImageIdx).Palette(4 * 2 ^ .BPP - 1)
                Case Else
                    ReDim G_ColorInfo(ImageIdx).Palette(1023)
            End Select
            
            '-- Read color info
            Get #hFile, , G_ColorInfo(ImageIdx).Palette()
            Get #hFile, , G_ColorInfo(ImageIdx).ColorIdxA
            Get #hFile, , G_ColorInfo(ImageIdx).ColorIdxB
            
            '-- Assign XOR palette [?]
            If (.BPP <= 8) Then
                Call G_oICON.oXORDIB(ImageIdx).SetPalette(G_ColorInfo(ImageIdx).Palette())
            End If
            '-- Assign AND palette
            Call G_oICON.oANDDIB(ImageIdx).SetPalette(aPalAND())
        End With
        
        '-- Read and assign XOR data
        ReDim aData(G_oICON.oXORDIB(ImageIdx).Size - 1)
        Get #hFile, , aData()
        Call CopyMemory(ByVal G_oICON.oXORDIB(ImageIdx).lpBits, aData(0), G_oICON.oXORDIB(ImageIdx).Size)
        
        '-- Read and assign AND data
        ReDim aData(G_oICON.oANDDIB(ImageIdx).Size - 1)
        Get #hFile, , aData()
        Call CopyMemory(ByVal G_oICON.oANDDIB(ImageIdx).lpBits, aData(0), G_oICON.oANDDIB(ImageIdx).Size)
    
    Close #hFile
End Sub

Private Sub pvRotateUndoFiles(ByVal ImageIdx As Integer)

  Dim lIdx As Long

    On Error Resume Next
    
    '-- Kill first
    Call Kill(pvPathItem(ImageIdx, 0))
    
    '-- Rotate: move up 1
    For lIdx = 1 To m_UNDOLEVELS
        Name pvPathItem(ImageIdx, lIdx - 0) As pvPathItem(ImageIdx, lIdx - 1)
    Next lIdx
    
    '-- Can not restore to first state
    m_uInfo(ImageIdx).Irreversible = True
    
    On Error GoTo 0
End Sub

Private Function pvPathList() As String

    '-- Build list files path
    pvPathList = m_sTemp & "\iICO" & m_AppID & "*.tmp"
End Function

Private Function pvPathItem(ByVal ImageIdx As Integer, ByVal UndoPos As Integer) As String

  Dim sFilter As String
    
    '-- History filter
    sFilter = IIf(UndoPos > -1, Format$(UndoPos, "000"), "*")
    '-- Build item file/s path
    pvPathItem = m_sTemp & "\iICO" & m_AppID & Format$(ImageIdx, "000") & sFilter & ".tmp"
End Function
