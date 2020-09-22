VERSION 5.00
Begin VB.Form fImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import "
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
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
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDiffuseError 
      Caption         =   "&Diffuse error"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3495
      TabIndex        =   9
      Top             =   2205
      Width           =   1350
   End
   Begin VB.CheckBox chkImport1bit 
      Caption         =   "Import &1-bit"
      Height          =   255
      Left            =   3495
      TabIndex        =   8
      Top             =   1905
      Width           =   1350
   End
   Begin VB.CheckBox chkImport8bit 
      Caption         =   "Import &8-bit"
      Height          =   255
      Left            =   3495
      TabIndex        =   6
      Top             =   1305
      Width           =   1350
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3495
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3585
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   4635
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3585
      Width           =   1050
   End
   Begin VB.Frame fraSize 
      Caption         =   "ARGB format/s"
      Height          =   1095
      Left            =   3495
      TabIndex        =   1
      Top             =   105
      Width           =   2190
      Begin VB.CheckBox chkSize 
         Caption         =   "&16 x 16"
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Tag             =   "16"
         Top             =   315
         Width           =   945
      End
      Begin VB.CheckBox chkSize 
         Caption         =   "&32 x 32"
         Height          =   285
         Index           =   2
         Left            =   1170
         TabIndex        =   4
         Tag             =   "32"
         Top             =   315
         Width           =   915
      End
      Begin VB.CheckBox chkSize 
         Caption         =   "&48 x 48"
         Height          =   285
         Index           =   3
         Left            =   1170
         TabIndex        =   5
         Tag             =   "48"
         Top             =   660
         Width           =   915
      End
      Begin VB.CheckBox chkSize 
         Caption         =   "&24 x 24"
         Height          =   285
         Index           =   1
         Left            =   210
         TabIndex        =   3
         Tag             =   "24"
         Top             =   660
         Width           =   900
      End
   End
   Begin VB.CheckBox chkResample 
      Caption         =   "&Resample"
      Height          =   255
      Left            =   3495
      TabIndex        =   10
      Top             =   2700
      Width           =   1350
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      Height          =   3030
      Left            =   180
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   195
      Width           =   3030
   End
   Begin VB.CheckBox chkImport4bit 
      Caption         =   "Import &4-bit"
      Height          =   255
      Left            =   3495
      TabIndex        =   7
      Top             =   1605
      Width           =   1350
   End
   Begin VB.CheckBox chkStretch 
      Caption         =   "&Stretch"
      Height          =   255
      Left            =   3495
      TabIndex        =   11
      Top             =   3000
      Width           =   1350
   End
   Begin VB.Line lnSep 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   233
      X2              =   379
      Y1              =   172
      Y2              =   172
   End
   Begin VB.Line lnSep 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   233
      X2              =   379
      Y1              =   171
      Y2              =   171
   End
End
Attribute VB_Name = "fImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Private variables:
Private m_oDIB32   As New cDIB
Private m_oTile    As New cTile
Private m_Pow2(31) As Long
Private m_Cancel   As Boolean

'//

Public Function Import(ByVal sFilename As String) As Boolean
    
  Const EDGEW As Long = 2
  
  Dim hBitmap As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  
  Dim bfx As Long, bfy As Long
  Dim bfW As Long, bfH As Long
  
    If (Len(sFilename)) Then
    
        Screen.MousePointer = vbHourglass
        
        '-- Load image...
        If (GdipLoadImageFromFile(StrConv(sFilename, vbUnicode), hBitmap) = [OK]) Then
            
            '-- Get dimensions
            Call GdipGetImageWidth(hBitmap, lWidth)
            Call GdipGetImageHeight(hBitmap, lHeight)
            
            '-- Scale/Convert to GDI DIB (preview size 200x200 -EDGEW)
            With picPreview
                Call m_oDIB32.GetBestFitInfo(lWidth, lHeight, .ScaleWidth - 2 * EDGEW, .ScaleHeight - 2 * EDGEW, bfx, bfy, bfW, bfH)
            End With
            Call m_oDIB32.Create(bfW, bfH, [32_bpp])
            Call BitmapToGDIDIB32(m_oDIB32, hBitmap)
            
            '-- Free image
            Call mGDIp.GdipDisposeImage(hBitmap)
        
            '-- Preview
            With picPreview
                Call .Cls
                Call m_oTile.Tile(.hDC, bfx + EDGEW, bfy + EDGEW, bfW, bfH)
                Call m_oDIB32.Stretch32(.hDC, bfx + EDGEW, bfy + EDGEW, bfW, bfH, , , , , , , Interpolate:=False)
            End With
            
            '-- Success
            Import = True
        End If
        
        Screen.MousePointer = vbDefault
    End If
End Function

Private Sub Form_Load()
    
  Dim lIdx As Long
    
    Set Me.Icon = Nothing
    
    Call mMisc.RemoveButtonBorderEnhance(Me.cmdOk)
    Call mMisc.RemoveButtonBorderEnhance(Me.cmdCancel)
    Call mMisc.ChangeBorderStyle(Me.picPreview.hWnd, [bsThin])
    
    '-- Load transparent layer pattern
    Call m_oTile.CreatePatternFromStdPicture(LoadResPicture("PATTERN_8X8", vbResBitmap))

    '-- Quick 2^n
    For lIdx = 0 To 30
        m_Pow2(lIdx) = 2 ^ lIdx
    Next lIdx
    m_Pow2(31) = &H80000000
End Sub

Private Sub chkImport8bit_Click()
    chkDiffuseError.Enabled = chkImport8bit Or chkImport8bit Or chkImport1bit
End Sub
Private Sub chkImport4bit_Click()
    chkDiffuseError.Enabled = chkImport8bit Or chkImport8bit Or chkImport1bit
End Sub
Private Sub chkImport1bit_Click()
    chkDiffuseError.Enabled = chkImport8bit Or chkImport8bit Or chkImport1bit
End Sub

'//

Private Sub cmdOk_Click()
    
  Dim bProcess   As Boolean
  
  Dim nIdx       As Integer
  Dim nImageIdx  As Integer
  
  Dim lSize      As Long
  Dim hBitmapSrc As Long
  Dim hBitmap    As Long
  Dim hGraphics  As Long
  
  Dim bfx As Long, bfy As Long
  Dim bfW As Long, bfH As Long
    
    For nIdx = 0 To chkSize.Count - 1
        bProcess = bProcess Or -chkSize(nIdx)
    Next nIdx
    
    If (Not bProcess) Then
        
        Call MsgBox("Choose import format/s please.", vbInformation)
        
      Else
      
        Screen.MousePointer = vbHourglass
        
        '-- New icon
        Call G_oICON.Destroy
        G_oICON.ResourceType = [rtIcon]
        
        nIdx = 0
        Do
            If (chkSize(nIdx)) Then
                
                '-- Get size
                lSize = chkSize(nIdx).Tag
                
                '-- Add format / get index
                nImageIdx = G_oICON.AddFormat(lSize, lSize, [ARGB_Color])
                
                '-- Build GDI+ bitmaps
                Call ARGBBitmapFromGDIDIB32(m_oDIB32, hBitmapSrc)
                Call ARGBBitmapFromGDIDIB32(G_oICON.oXORDIB(nImageIdx), hBitmap)
                Call GdipGetImageGraphicsContext(hBitmap, hGraphics)
                
                '-- Set render props.
                Call GdipSetInterpolationMode(hGraphics, [InterpolationModeNearestNeighbor] + (2 * chkResample))
                Call GdipSetPixelOffsetMode(hGraphics, [PixelOffsetModeHighQuality])
                If (chkStretch) Then
                    bfx = 0: bfW = lSize
                    bfy = 0: bfH = lSize
                  Else
                    Call m_oDIB32.GetBestFitInfo(m_oDIB32.Width, m_oDIB32.Height, lSize, lSize, bfx, bfy, bfW, bfH, StretchFit:=True)
                End If
                
                '-- Scale and back to GDI DIB
                Call GdipDrawImageRectRectI(hGraphics, hBitmapSrc, bfx, bfy, bfW, bfH, 0, 0, m_oDIB32.Width, m_oDIB32.Height, UnitPixel, 0, 0, 0)
                Call ARGBBitmapToGDIDIB32(G_oICON.oXORDIB(nImageIdx), hBitmap)
                
                '-- Clean up GDI+
                Call GdipDisposeImage(hBitmap)
                Call GdipDisposeImage(hBitmapSrc)
                Call GdipDeleteGraphics(hGraphics)
                
            End If
            nIdx = nIdx + 1
        Loop Until nIdx = chkSize.Count
        
        If (chkImport1bit Or chkImport4bit Or chkImport8bit) Then
        
            nIdx = 0
            Do
                If (chkSize(nIdx)) Then
                    
                    '-- Get size
                    lSize = chkSize(nIdx).Tag
                    
                    With G_oICON
                        
                        If (chkImport1bit) Then
                        
                            '-- Add format / get index
                            nImageIdx = .AddFormat(lSize, lSize, [002_Colors])
                            
                            '-- Extract 8bit from ARGB format
                            Call pvExtract1bitFormat( _
                                 .oXORDIB(nImageIdx), _
                                 .oANDDIB(nImageIdx), _
                                 .oXORDIB(.GetFormatIndex(lSize, lSize, [ARGB_Color])))
                        End If
                        
                        If (chkImport4bit) Then
                        
                            '-- Add format / get index
                            nImageIdx = .AddFormat(lSize, lSize, [016_Colors])
                            
                            '-- Extract 8bit from ARGB format
                            Call pvExtract4bitFormat( _
                                 .oXORDIB(nImageIdx), _
                                 .oANDDIB(nImageIdx), _
                                 .oXORDIB(.GetFormatIndex(lSize, lSize, [ARGB_Color])))
                        End If
                        
                        If (chkImport8bit) Then
                            
                            '-- Add format / get index
                            nImageIdx = .AddFormat(lSize, lSize, [256_Colors])
                            
                            '-- Extract 8bit from ARGB format
                            Call pvExtract8bitFormat( _
                                 .oXORDIB(nImageIdx), _
                                 .oANDDIB(nImageIdx), _
                                 .oXORDIB(.GetFormatIndex(lSize, lSize, [ARGB_Color])))
                        End If
                        
                    End With
                End If
                nIdx = nIdx + 1
            Loop Until nIdx = chkSize.Count
        End If
        
        Screen.MousePointer = vbDefault
        
        m_Cancel = False
        Call m_oDIB32.Destroy
        Call Me.Hide
    End If
End Sub

Private Sub cmdCancel_Click()
    m_Cancel = True
    Call m_oDIB32.Destroy
    Call Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '-- Close button pressed
    If (UnloadMode = vbFormControlMenu) Then
        Call cmdCancel_Click
    End If
End Sub

'//

Public Property Get Cancel() As Boolean
    Cancel = m_Cancel
End Property

'//

Private Sub pvExtract1bitFormat(oXORDIB As cDIB, oANDDIB As cDIB, oARGBDIB As cDIB)
    
  Dim oDither    As New cDIBDither
  Dim oDIB32     As New cDIB
  Dim oPal       As New cPalette
  Dim oICEx      As New cIconCanvasEx
  Dim aPal(7)    As Byte
    
    '-- Resize ARGB source (temp) and blend over white surface
    Call oDIB32.Create(oARGBDIB.Width, oARGBDIB.Height, [32_bpp])
    Call oDIB32.Cls(vbWhite)
    Call oARGBDIB.Stretch32(oDIB32.hDC, 0, 0, oARGBDIB.Width, oARGBDIB.Height)
    
    '-- Create monochrome palette
    Call oPal.CreateGreyScale([002_pgColors])
    
    '-- XOR: Dither to new palette
    Call oDither.DitherToGreyScale(oDIB32, oXORDIB, Diffuse:=chkDiffuseError)
    Call CopyMemory(aPal(0), ByVal oPal.lpPalette, 7)
    Call oXORDIB.SetPalette(aPal())
    '-- AND: Reset (fill with black)
    Call oANDDIB.Cls(vbBlack)
    
    '-- Process solid mask
    Call oICEx.ProcessSolidIconFromAlphaDIB(oXORDIB, oANDDIB, oARGBDIB)
End Sub

Private Sub pvExtract4bitFormat(oXORDIB As cDIB, oANDDIB As cDIB, oARGBDIB As cDIB)
    
  Dim oDither    As New cDIBDither
  Dim oDIB32     As New cDIB
  Dim oPal       As New cPalette
  Dim oICEx      As New cIconCanvasEx
  Dim aPal(63)   As Byte
    
    '-- Resize ARGB source (temp) and blend over white surface
    Call oDIB32.Create(oARGBDIB.Width, oARGBDIB.Height, [32_bpp])
    Call oDIB32.Cls(vbWhite)
    Call oARGBDIB.Stretch32(oDIB32.hDC, 0, 0, oARGBDIB.Width, oARGBDIB.Height)
    
    '-- Create EGA palette
    Call oPal.CreateEGA
    
    '-- XOR: Dither to new palette
    Call oDither.DitherToColorPalette(oPal, oDIB32, oXORDIB, Diffuse:=chkDiffuseError)
    Call CopyMemory(aPal(0), ByVal oPal.lpPalette, 63)
    Call oXORDIB.SetPalette(aPal())
    '-- AND: Reset (fill with black)
    Call oANDDIB.Cls(vbBlack)
    
    '-- Process solid mask
    Call oICEx.ProcessSolidIconFromAlphaDIB(oXORDIB, oANDDIB, oARGBDIB)
End Sub

Private Sub pvExtract8bitFormat(oXORDIB As cDIB, oANDDIB As cDIB, oARGBDIB As cDIB)
    
  Dim oDither    As New cDIBDither
  Dim oDIB32     As New cDIB
  Dim oPal       As New cPalette
  Dim oICEx      As New cIconCanvasEx
  Dim aPal(1023) As Byte
    
    '-- Resize ARGB source (temp) and blend over white surface
    Call oDIB32.Create(oARGBDIB.Width, oARGBDIB.Height, [32_bpp])
    Call oDIB32.Cls(vbWhite)
    Call oARGBDIB.Stretch32(oDIB32.hDC, 0, 0, oARGBDIB.Width, oARGBDIB.Height)
    
    '-- Create optimal palette
    Call oPal.CreateOptimal(oDIB32, 255, 8): oPal.Entries = 256
    Call oPal.SortPalette
    Call oPal.BuildLogicalPalette
    
    '-- XOR: Dither to new palette
    Call oDither.DitherToColorPalette(oPal, oDIB32, oXORDIB, Diffuse:=chkDiffuseError)
    Call CopyMemory(aPal(0), ByVal oPal.lpPalette, 1024)
    Call oXORDIB.SetPalette(aPal())
    '-- AND: Reset (fill with black)
    Call oANDDIB.Cls(vbBlack)
    
    '-- Process solid mask
    Call oICEx.ProcessSolidIconFromAlphaDIB(oXORDIB, oANDDIB, oARGBDIB)
End Sub
