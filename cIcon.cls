VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cIcon.cls
' Author:        Carles P.V.
' Dependencies:  cDIB.cls, cPalette.cls
' Last revision: 2004.06.14
'================================================

Option Explicit

'-- API:

Private Type ICONDIR
    idReserved    As Integer  ' Reserved
    idType        As Integer  ' Resource type
    idCount       As Integer  ' Image Count
End Type '6 bytes

Private Type ICONDIRENTRY
    bWidth        As Byte     ' Width of the image
    bHeight       As Byte     ' Height of the image (2 * Height)
    bColorCount   As Byte     ' Number of colors in image (0 when >= 8 bpp)
    bReserved     As Byte     ' Reserved
    wPlanes       As Integer  ' Color Planes   (-> xHotspot [Cursor])
    wBitCount     As Integer  ' Bits per pixel (-> yHotspot [Cursor])
    dwBytesInRes  As Long     ' How many bytes in this resource?
    dwImageOffset As Long     ' Where in the file is this image?
End Type '16 bytes

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type '40 bytes

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDst As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)

'//

'-- Public Enums.:

Public Enum ResourceTypeCts
    [rtIcon] = 1
    [rtCursor] = 2
End Enum

Public Enum icoBPPCts
    [002_Colors] = 1
    [016_Colors] = 4
    [256_Colors] = 8
    [True_Color] = 24
    [ARGB_Color] = 32
End Enum

'-- Private Types:
Private Type DIBDATA
    XORDIB As cDIB                   ' XOR DIB section
    ANDDIB As cDIB                   ' AND DIB section
End Type

'-- Private Variables:
Private m_Loaded     As Boolean      ' Icon <loaded> flag
Private m_Dir        As ICONDIR      ' Icon file header
Private m_DirEntry() As ICONDIRENTRY ' Icon image headers
Private m_Data()     As DIBDATA      ' Icon data (DIBs)
Private m_OrderKey() As String * 8   ' Image format key

 

'========================================================================================
' Class
'========================================================================================

Private Sub Class_Initialize()
    m_Dir.idType = [rtIcon]
End Sub

Private Sub Class_Terminate()
    Call pvClear
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub Destroy()
    Call pvClear
End Sub

Public Function LoadFromFile(ByVal Filename As String, Optional ByVal SortByFormat As Boolean = True) As Boolean
    
  Dim currRes    As ResourceTypeCts
  Dim uBIH       As BITMAPINFOHEADER
  Dim aXORBits() As Byte
  Dim aANDBits() As Byte
  Dim aPalXOR()  As Byte
  Dim aPalAND(7) As Byte

  Dim hFile As Integer
  Dim nImg  As Integer
    
    '-- Store temp. resource type (->Err)
    currRes = m_Dir.idType
    
    '-- Clear / Activate error handling
    Call pvClear
    On Error GoTo ErrLoad
    
    '-- Open file
    hFile = FreeFile()
    Open Filename For Binary Access Read As hFile
      
    '-- Get icon header
    Get #hFile, , m_Dir
    '-- Get icon entries
    ReDim m_DirEntry(m_Dir.idCount - 1)
    Get #hFile, , m_DirEntry()
    
    '-- Initialize arrays and monochrome palette
    ReDim m_OrderKey(m_Dir.idCount - 1)
    ReDim m_Data(m_Dir.idCount - 1)
    ReDim m_DIBData(m_Dir.idCount - 1)
    Call FillMemory(aPalAND(4), 3, &HFF)
    
    '-- Get images
    For nImg = 0 To m_Dir.idCount - 1
        
        '-- Move to begin of image data
        Seek #hFile, m_DirEntry(nImg).dwImageOffset + 1
        
        '-- Load BITMAPINFOHEADER
        Get #hFile, , uBIH
        
        '-- Load XOR palette [?] (<= 8 bpp)
        If (uBIH.biBitCount <= 8) Then
            ReDim aPalXOR(4 * 2 ^ uBIH.biBitCount - 1): Get #hFile, , aPalXOR()
        End If
        
        With m_Data(nImg)
            
            '-- Inititalize XOR DIB
            Set .XORDIB = New cDIB
            Call .XORDIB.Create(uBIH.biWidth, uBIH.biHeight \ 2, uBIH.biBitCount)
            If (uBIH.biBitCount <= 8) Then
                Call .XORDIB.SetPalette(aPalXOR())
            End If
            
            '-- Inititalize AND DIB
            Set .ANDDIB = New cDIB
            Call .ANDDIB.Create(uBIH.biWidth, uBIH.biHeight \ 2, [01_bpp])
            Call .ANDDIB.SetPalette(aPalAND())
            
            '-- Read DIB bits
            ReDim aXORBits(.XORDIB.Size - 1): Get #hFile, , aXORBits()
            ReDim aANDBits(.ANDDIB.Size - 1): Get #hFile, , aANDBits()
            
            '-- Assign DIB bits
            Call CopyMemory(ByVal .XORDIB.lpBits, aXORBits(0), .XORDIB.Size)
            Call CopyMemory(ByVal .ANDDIB.lpBits, aANDBits(0), .ANDDIB.Size)
            
            '-- Build image format key (sort entries)
            m_OrderKey(nImg) = Format$(uBIH.biWidth, "000") & Format$(uBIH.biHeight \ 2, "000") & Format$(uBIH.biBitCount, "00")
        End With
    Next nImg

    Close #hFile
    If (SortByFormat) Then
        Call pvSortEntries
    End If
    
'-- Success
    m_Loaded = True
    LoadFromFile = True
    Exit Function
    
'-- Err.
ErrLoad:
    Close #hFile
    Call pvClear
    m_Dir.idType = currRes
End Function

Public Function SaveToFile(ByVal Filename As String) As Boolean

  Dim uBIH       As BITMAPINFOHEADER
  Dim aPalXOR()  As Byte
  Dim aXORBits() As Byte
  Dim aANDBits() As Byte

  Dim hFile As Integer
  Dim nImg  As Integer


    If (m_Loaded) Then
    
        On Error Resume Next
        Call Kill(Filename)
        On Error GoTo 0

        On Error GoTo ErrSave
    
        hFile = FreeFile()
        Open Filename For Binary Access Write As hFile

        '-- Write icon header
        Put #hFile, , m_Dir
        '-- Write icon entries
        Put #hFile, , m_DirEntry()

        '-- Write icon data
        For nImg = 0 To m_Dir.idCount - 1

            With m_Data(nImg)
            
                '-- Build BITMAPINFOHEADER / Get palette [?]
                With .XORDIB
                    uBIH.biSize = Len(uBIH)
                    uBIH.biPlanes = 1
                    uBIH.biBitCount = .BPP
                    uBIH.biWidth = .Width
                    uBIH.biHeight = 2 * .Height
                    If (.BPP <= 8) Then
                        ReDim aPalXOR(4 * 2 ^ .BPP - 1): Call .GetPalette(aPalXOR())
                    End If
                End With

                '-- Get DIB bits
                ReDim aXORBits(.XORDIB.Size - 1)
                ReDim aANDBits(.ANDDIB.Size - 1)
                Call CopyMemory(aXORBits(0), ByVal .XORDIB.lpBits, .XORDIB.Size)
                Call CopyMemory(aANDBits(0), ByVal .ANDDIB.lpBits, .ANDDIB.Size)
    
                '-- Move to entry start
                Seek #hFile, m_DirEntry(nImg).dwImageOffset + 1
                
                '-- Write data
                Put #hFile, , uBIH
                If (.XORDIB.BPP <= 8) Then
                    Put #hFile, , aPalXOR()
                End If
                Put #hFile, , aXORBits()
                Put #hFile, , aANDBits()
            End With
        Next nImg

        Close #hFile
        SaveToFile = True
    End If
    Exit Function

ErrSave:
    Close #hFile
End Function

'//

Public Function AddFormat(ByVal NewWidth As Integer, ByVal NewHeight As Integer, ByVal NewBPP As icoBPPCts, Optional ByVal SortByFormat As Boolean = True) As Integer

  Dim nCount As Integer
  Dim oPal   As New cPalette
  Dim aPal() As Byte

    '-- Check format
    If (Not pvFormatExists(Format$(NewWidth, "000") & Format$(NewHeight, "000") & Format$(NewBPP, "00"))) Then

        '-- Increase image Count
        nCount = m_Dir.idCount
        m_Dir.idCount = m_Dir.idCount + 1

        '-- Redim. arrays
        ReDim Preserve m_OrderKey(nCount)
        ReDim Preserve m_DirEntry(nCount)
        ReDim Preserve m_Data(nCount)

        '-- Define icon entry
        With m_DirEntry(nCount)
            .wPlanes = 1
            .wBitCount = NewBPP
            .bWidth = NewWidth
            .bHeight = NewHeight
            .dwBytesInRes = 40 + IIf(NewBPP <= 8, 4 * 2 ^ NewBPP, 0) + 4 * ((NewWidth * NewBPP + 31) \ 32) * NewHeight + 4 * ((NewWidth * 1 + 31) \ 32) * NewHeight
        End With

        '-- Build icon (XOR and AND DIBs)
        With m_Data(nCount)
        
            '-- Initialize XOR DIB
            Set .XORDIB = New cDIB
            
            With .XORDIB
                
                '-- Create XOR DIB (Use ImageXORDIB public object to set/initialize palette)
                Call .Create(NewWidth, NewHeight, NewBPP)
                
                '-- Initialize palette [?]
                If (NewBPP <= 8) Then
                    Select Case NewBPP
                        Case [002_Colors]
                            Call oPal.CreateGreyScale([002_pgColors])
                        Case [016_Colors]
                            Call oPal.CreateEGA
                        Case [256_Colors]
                            Call oPal.CreateSpectrum
                    End Select
                    ReDim aPal(0 To 4 * oPal.Entries - 1)
                    Call CopyMemory(aPal(0), ByVal oPal.lpPalette, 4 * oPal.Entries)
                    Call .SetPalette(aPal())
                End If
                
                '-- Reset DIB bits: W9x
                Call .Cls(&H0)
            End With

            '-- Initialize AND DIB
            Set .ANDDIB = New cDIB
            
            With .ANDDIB
                
                '-- Initialize palette
                ReDim aPal(0 To 7)
                Call oPal.CreateGreyScale([002_pgColors])
                Call CopyMemory(aPal(0), ByVal oPal.lpPalette, 8)
                
                '-- Create AND DIB
                Call .Create(NewWidth, NewHeight, [01_bpp])
                Call .SetPalette(aPal())
                
                '-- Reset DIB bits: W9x
                Call .Cls(&HFFFFFF)
            End With
            
            '-- Cursor [?]
            If (m_Dir.idType = [rtCursor]) Then
                m_DirEntry(nCount).wPlanes = 0   '-> Hot spot X
                m_DirEntry(nCount).wBitCount = 0 '-> Hot spot y
            End If
        End With

        '-- Build sort key and sort
        m_OrderKey(nCount) = Format$(NewWidth, "000") & Format$(NewHeight, "000") & Format$(NewBPP, "00")
        If (SortByFormat) Then
            Call pvSortEntries
        End If
        
        '-- Set <loaded> flag to true (case 1st entry)
        m_Loaded = True
        AddFormat = Me.GetFormatIndex(NewWidth, NewHeight, NewBPP)
      Else
        AddFormat = -1
    End If
End Function

Public Function RemoveFormat(ByVal nIcon As Integer) As Boolean

  Dim nCount As Integer
    
    '-- Current Count
    nCount = m_Dir.idCount
    
    '-- More than one entry [?]
    If (nCount > 1) Then
    
        '-- Set key as 'last' key
        m_OrderKey(nIcon) = "99999999"
        '-- Sort entries
        Call pvSortEntries
        
        '-- Remove last
        m_Dir.idCount = nCount - 1
        Set m_Data(m_Dir.idCount).XORDIB = Nothing
        Set m_Data(m_Dir.idCount).ANDDIB = Nothing
        ReDim Preserve m_OrderKey(m_Dir.idCount)
        ReDim Preserve m_DirEntry(m_Dir.idCount)
        ReDim Preserve m_Data(m_Dir.idCount)
        
        '-- Success
        RemoveFormat = True
    End If
End Function

'//

Public Function DrawIconFit(ByVal nIcon As Integer, ByVal hDC As Long, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal Fit As Boolean = False, Optional ByVal WidthFit As Long = 0, Optional ByVal HeightFit As Long = 0) As Boolean

  Dim W As Long, cW As Single, bfx As Long, bfW As Long
  Dim H As Long, cH As Single, bfy As Long, bfH As Long
    
    If (m_Loaded) Then
        
        If (nIcon >= 0 And nIcon < m_Dir.idCount) Then
        
            W = m_Data(nIcon).XORDIB.Width
            H = m_Data(nIcon).XORDIB.Height
                        
            '-- Get best fit dimensions/offsets [?]
            If (Fit) Then
                If (W > WidthFit Or H > HeightFit) Then
                    cW = WidthFit / W
                    cH = HeightFit / H
                    If (cW < cH) Then
                        bfW = WidthFit
                        bfH = W * cW
                      Else
                        bfH = HeightFit
                        bfW = H * cH
                    End If
                  Else
                    bfW = W
                    bfH = H
                End If
                '-- Get x and y offsets
                bfx = (WidthFit - bfW) \ 2
                bfy = (HeightFit - bfH) \ 2
              Else
                bfW = W
                bfH = H
            End If
            
            '-- Paint it
            With m_Data(nIcon)
                
                If (.XORDIB.BPP = [32_bpp]) Then
                    '-- ARGB format
                    Call .XORDIB.Stretch32(hDC, x + bfx, y + bfy, bfW, bfH)
                  Else
                    '-- Any other format
                    Call .ANDDIB.Stretch(hDC, x + bfx, y + bfy, bfW, bfH, , , , , vbSrcAnd)
                    Call .XORDIB.Stretch(hDC, x + bfx, y + bfy, bfW, bfH, , , , , vbSrcPaint)
                End If
            End With
            DrawIconFit = True
        End If
    End If
End Function

Public Function DrawIconStretch(ByVal nIcon As Integer, ByVal hDC As Long, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ScaleFactor As Single = 1) As Boolean

  Dim W As Long
  Dim H As Long
    
    If (m_Loaded) Then
        
        If (nIcon >= 0 And nIcon < m_Dir.idCount) Then
        
            W = ScaleFactor * m_Data(nIcon).XORDIB.Width
            H = ScaleFactor * m_Data(nIcon).XORDIB.Height
                        
            '-- Paint it
            With m_Data(nIcon)
                
                If (.XORDIB.BPP = [32_bpp]) Then
                    '-- ARGB format
                    Call .XORDIB.Stretch32(hDC, x, y, W, H)
                  Else
                    '-- Any other format
                    Call .ANDDIB.Stretch(hDC, x, y, W, H, , , , , vbSrcAnd)
                    Call .XORDIB.Stretch(hDC, x, y, W, H, , , , , vbSrcPaint)
                End If
            End With
            DrawIconStretch = True
        End If
    End If
End Function

'//

Public Sub SortByFormat()

    '-- Sort entries ###[W]###[H]###[BPP]
    Call pvSortEntries
End Sub

Public Function GetFormatIndex(ByVal Width As Integer, ByVal Height As Integer, ByVal BPP As icoBPPCts) As Integer

  Dim sKey As String
  Dim nImg As Integer
    
    '-- Get icon index from given format key
    sKey = Format$(Width, "000") & Format$(Height, "000") & Format$(BPP, "00")
    
    For nImg = 0 To m_Dir.idCount - 1
        If (sKey = m_OrderKey(nImg)) Then
            GetFormatIndex = nImg
            Exit Function
        End If
    Next nImg
    
    GetFormatIndex = -1
End Function

'========================================================================================
' Properties
'========================================================================================

'-- Use it carefully!
Public Property Get oXORDIB(ByVal nIcon As Integer) As cDIB
    Set oXORDIB = m_Data(nIcon).XORDIB
End Property

'-- Use it carefully!
Public Property Get oANDDIB(ByVal nIcon As Integer) As cDIB
    Set oANDDIB = m_Data(nIcon).ANDDIB
End Property

'//

Public Property Get ResourceType() As ResourceTypeCts
    ResourceType = m_Dir.idType
End Property
Public Property Let ResourceType(ByVal ID As ResourceTypeCts)
    m_Dir.idType = ID
End Property

Public Property Get Count() As Integer
    If (m_Loaded) Then
        Count = m_Dir.idCount
    End If
End Property

Public Property Get Width(ByVal nIcon As Integer) As Integer
    If (m_Loaded) Then
        Width = m_Data(nIcon).XORDIB.Width
    End If
End Property

Public Property Get Height(ByVal nIcon As Integer) As Integer
    If (m_Loaded) Then
        Height = m_Data(nIcon).XORDIB.Height
    End If
End Property

Public Property Get BPP(ByVal nIcon As Integer) As icoBPPCts
    If (m_Loaded) Then
        BPP = m_Data(nIcon).XORDIB.BPP
    End If
End Property

Public Property Get ImageSize(ByVal nIcon As Integer) As Long
    If (m_Loaded) Then
        ImageSize = m_DirEntry(nIcon).dwBytesInRes
    End If
End Property

Public Property Get ImagesSize() As Long

  Dim nImg As Integer
  
    If (m_Loaded) Then
        For nImg = 0 To m_Dir.idCount - 1
            ImagesSize = ImagesSize + m_DirEntry(nImg).dwBytesInRes
        Next nImg
    End If
End Property

'//

Public Property Get HotSpotX(ByVal nIcon As Integer) As Integer
    If (m_Loaded) Then
        HotSpotX = m_DirEntry(nIcon).wPlanes
    End If
End Property
Public Property Let HotSpotX(ByVal nIcon As Integer, ByVal New_HotSpotX As Integer)
    If (m_Loaded) Then
        m_DirEntry(nIcon).wPlanes = New_HotSpotX
    End If
End Property

Public Property Get HotSpotY(ByVal nIcon As Integer) As Integer
    If (m_Loaded) Then
        HotSpotY = m_DirEntry(nIcon).wBitCount
    End If
End Property
Public Property Let HotSpotY(ByVal nIcon As Integer, ByVal New_HotSpotY As Integer)
    If (m_Loaded) Then
        m_DirEntry(nIcon).wBitCount = New_HotSpotY
    End If
End Property

'========================================================================================
' Private
'========================================================================================

Private Sub pvClear()

  Dim nImg As Integer
  
    '-- Reset header/descriptors
    m_Dir.idCount = 0
    Erase m_DirEntry()
    Erase m_OrderKey()
    
    '-- Destroy DIB sections
    If (m_Loaded) Then
        For nImg = 0 To m_Dir.idCount - 1
            Set m_Data(nImg).XORDIB = Nothing
            Set m_Data(nImg).ANDDIB = Nothing
        Next nImg
        Erase m_Data()
    End If
    
    '-- Reset <loaded> flag
    m_Loaded = False
End Sub

Private Function pvFormatExists(ByVal FormatKey As String) As Boolean
    
  Dim nImg As Integer
    
    '-- Check format given the format key
    For nImg = 0 To m_Dir.idCount - 1
        If (FormatKey = m_OrderKey(nImg)) Then
            pvFormatExists = True
            Exit For
        End If
    Next nImg
End Function

Private Sub pvSortEntries()
    
  Const ICON_FILEHEADER  As Long = 6
  Const ICON_ENTRYHEADER As Long = 16
  
  Dim sOrderKey As String * 8
  Dim uDirEntry As ICONDIRENTRY
  Dim uEntry    As DIBDATA
  
  Dim n1   As Integer
  Dim n2   As Integer
  Dim nImg As Integer
    
    '-- Sort by Size, Color depth
    If (m_Dir.idCount > 1) Then
    
        For n1 = 0 To m_Dir.idCount - 2
        For n2 = n1 To m_Dir.idCount - 1
        
            If (m_OrderKey(n2) < m_OrderKey(n1)) Then
                
                '-- Swap key!
                sOrderKey = m_OrderKey(n1)
                m_OrderKey(n1) = m_OrderKey(n2)
                m_OrderKey(n2) = sOrderKey
                
                '-- Swap dir entry
                uDirEntry = m_DirEntry(n1)
                m_DirEntry(n1) = m_DirEntry(n2)
                m_DirEntry(n2) = uDirEntry
                
                '-- Swap icon data
                uEntry = m_Data(n1)
                m_Data(n1) = m_Data(n2)
                m_Data(n2) = uEntry
            End If
        Next n2, n1
    End If
    
    '-- Update image offsets
    m_DirEntry(0).dwImageOffset = ICON_FILEHEADER + m_Dir.idCount * ICON_ENTRYHEADER
    
    For nImg = 1 To m_Dir.idCount - 1
        m_DirEntry(nImg).dwImageOffset = m_DirEntry(nImg - 1).dwImageOffset + m_DirEntry(nImg - 1).dwBytesInRes
    Next nImg
End Sub

