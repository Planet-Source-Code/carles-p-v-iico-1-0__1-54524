VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "iICO"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8610
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   574
   Begin iICO.ucIconInfo ucIconInfo 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      Top             =   6765
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   476
   End
   Begin VB.Frame fraDrawTools 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   3975
      TabIndex        =   2
      Top             =   1125
      Visible         =   0   'False
      Width           =   4725
      Begin VB.CheckBox chkImport32bppOpaque 
         Caption         =   "Opaque"
         Height          =   210
         Left            =   1725
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   60
         Width           =   885
      End
      Begin VB.CheckBox chkExportScreenAsColorB 
         Caption         =   "Export Screen as B"
         Height          =   210
         Left            =   2850
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   60
         Width           =   2145
      End
      Begin VB.CheckBox chkImportColorBAsScreen 
         Caption         =   "Import B as Screen"
         Height          =   210
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   60
         Width           =   1665
      End
   End
   Begin VB.Frame fraDrawTools 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   3975
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   4725
      Begin iICO.ucToolbar ucToolbarPixelMask 
         Height          =   270
         Left            =   915
         Top             =   -15
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   476
         BarStyle        =   2
      End
      Begin VB.Label lblPixelMask 
         Caption         =   "Pixel mask"
         Height          =   255
         Left            =   15
         TabIndex        =   14
         Top             =   60
         Width           =   930
      End
   End
   Begin VB.PictureBox picPreview 
      ClipControls    =   0   'False
      Height          =   1980
      Left            =   240
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1980
   End
   Begin VB.Frame fraDrawTools 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   9
      Left            =   3975
      TabIndex        =   10
      Top             =   4470
      Visible         =   0   'False
      Width           =   4725
      Begin VB.Label lblHotSpotV 
         Height          =   255
         Left            =   1260
         TabIndex        =   33
         Top             =   75
         Width           =   1725
      End
      Begin VB.Label lblHotSpot 
         Caption         =   "Hot spot coord.:"
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   75
         Width           =   1275
      End
   End
   Begin VB.Frame fraDrawTools 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   8
      Left            =   3975
      TabIndex        =   9
      Top             =   4095
      Visible         =   0   'False
      Width           =   4725
      Begin VB.CheckBox chkPickAlpha 
         Caption         =   "Pick alpha"
         Height          =   210
         Left            =   0
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.Frame fraDrawTools 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   7
      Left            =   3975
      TabIndex        =   8
      Top             =   3735
      Visible         =   0   'False
      Width           =   4725
      Begin iICO.ucToolbar ucToolbarFont 
         Height          =   315
         Left            =   3870
         Top             =   -15
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   556
         BarStyle        =   2
      End
      Begin VB.TextBox txtText 
         Height          =   315
         Left            =   0
         MaxLength       =   25
         TabIndex        =   28
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox cbFont 
         Height          =   315
         Left            =   930
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   0
         Width           =   2100
      End
      Begin VB.ComboBox cbFontSize 
         Height          =   315
         Left            =   3090
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.Frame fraDrawTools 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   5
      Left            =   3975
      TabIndex        =   6
      Top             =   2865
      Visible         =   0   'False
      Width           =   4725
      Begin VB.ComboBox cbShapeLineWidth 
         Height          =   315
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   0
         Width           =   795
      End
      Begin VB.ComboBox cbShape 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   0
         Width           =   1680
      End
      Begin VB.Label lblShapeLineWidth 
         Caption         =   "Shape width"
         Height          =   255
         Left            =   2730
         TabIndex        =   21
         Top             =   60
         Width           =   885
      End
      Begin VB.Label lblShape 
         Caption         =   "Shape type"
         Height          =   255
         Left            =   15
         TabIndex        =   19
         Top             =   60
         Width           =   885
      End
   End
   Begin VB.Frame fraDrawTools 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   6
      Left            =   3975
      TabIndex        =   7
      Top             =   3300
      Visible         =   0   'False
      Width           =   4725
      Begin iICO.ucUpDownBox ucARGBTolerance 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   24
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Max             =   255
      End
      Begin iICO.ucUpDownBox ucARGBTolerance 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   25
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Max             =   360
      End
      Begin iICO.ucUpDownBox ucARGBTolerance 
         Height          =   315
         Index           =   2
         Left            =   2880
         TabIndex        =   26
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin iICO.ucUpDownBox ucARGBTolerance 
         Height          =   315
         Index           =   3
         Left            =   3720
         TabIndex        =   27
         Top             =   0
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTolerance 
         Caption         =   "AHSL tolerance"
         Height          =   255
         Left            =   15
         TabIndex        =   23
         Top             =   60
         Width           =   1245
      End
   End
   Begin VB.Frame fraDrawTools 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   3975
      TabIndex        =   4
      Top             =   1995
      Visible         =   0   'False
      Width           =   4725
      Begin VB.ComboBox cbStraightLineWidth 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   0
         Width           =   795
      End
      Begin VB.Label lblStraightLineWidth 
         Caption         =   "Line width"
         Height          =   255
         Left            =   15
         TabIndex        =   15
         Top             =   60
         Width           =   885
      End
   End
   Begin VB.Frame fraDrawTools 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   4
      Left            =   3975
      TabIndex        =   5
      Top             =   2430
      Visible         =   0   'False
      Width           =   4725
      Begin VB.ComboBox cbBrushLineWidth 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   0
         Width           =   795
      End
      Begin VB.Label lblBrushLineWidth 
         Caption         =   "Brush width"
         Height          =   255
         Left            =   15
         TabIndex        =   17
         Top             =   60
         Width           =   885
      End
   End
   Begin iICO.ucSlider ucAlphaPicker 
      Height          =   315
      Left            =   2520
      Top             =   1275
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
   End
   Begin iICO.ucPalettePicker ucPalettePicker 
      Height          =   4800
      Left            =   2505
      Top             =   1635
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   8467
   End
   Begin iICO.ucColor ucColorA 
      Height          =   390
      Left            =   3015
      ToolTipText     =   "Color A"
      Top             =   765
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   688
   End
   Begin iICO.ucColor ucColorB 
      Height          =   390
      Left            =   3465
      ToolTipText     =   "Color B"
      Top             =   765
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   688
   End
   Begin iICO.ucColor ucColorScreen 
      Height          =   390
      Left            =   2535
      ToolTipText     =   "Color Screen"
      Top             =   765
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   688
      PickColorCursor =   -1  'True
      IsColorScreen   =   -1  'True
   End
   Begin iICO.ucToolbar ucToolbarFormat 
      Height          =   495
      Left            =   240
      Top             =   600
      Width           =   1335
      _ExtentX        =   3413
      _ExtentY        =   873
      BarStyle        =   2
   End
   Begin iICO.ucToolbar ucToolbarDrawTools 
      Height          =   405
      Left            =   4125
      Top             =   750
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   714
      BarStyle        =   2
   End
   Begin iICO.ucToolbar ucToolbarMain 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   661
      BarStyle        =   2
      BarEdge         =   -1  'True
   End
   Begin iICO.ucIconList ucIconList 
      Height          =   3165
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1185
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   5583
      ThumbnailHeight =   64
   End
   Begin iICO.ucIconCanvas ucIconCanvas 
      Height          =   1455
      Left            =   3960
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4815
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2328
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New..."
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &as..."
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Import..."
         Index           =   6
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Explorer..."
         Index           =   7
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   9
      End
   End
   Begin VB.Menu mnuEditTop 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Undo"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Redo"
         Index           =   1
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cu&t"
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Copy"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Paste"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "C&lear"
         Index           =   6
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select &all"
         Index           =   8
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuPaletteTop 
      Caption         =   "&Palette"
      Begin VB.Menu mnuPalette 
         Caption         =   "&Load..."
         Index           =   0
      End
      Begin VB.Menu mnuPalette 
         Caption         =   "&Save..."
         Index           =   1
      End
      Begin VB.Menu mnuPalette 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPalette 
         Caption         =   "&Predefined"
         Index           =   3
         Begin VB.Menu mnuPalettePredefined 
            Caption         =   "&Web safe"
            Index           =   0
         End
         Begin VB.Menu mnuPalettePredefined 
            Caption         =   "&Grey scale"
            Index           =   1
         End
         Begin VB.Menu mnuPalettePredefined 
            Caption         =   "&Spectrum"
            Index           =   2
         End
         Begin VB.Menu mnuPalettePredefined 
            Caption         =   "Windows &XP"
            Index           =   3
         End
         Begin VB.Menu mnuPalettePredefined 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuPalettePredefined 
            Caption         =   "&Optimal"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnuCanvasTop 
      Caption         =   "&Canvas"
      Begin VB.Menu mnuCanvas 
         Caption         =   "&Screen color..."
         Index           =   0
      End
      Begin VB.Menu mnuCanvas 
         Caption         =   "Pixel &grid"
         Index           =   1
      End
      Begin VB.Menu mnuCanvas 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCanvas 
         Caption         =   "&Antialias"
         Index           =   3
      End
   End
   Begin VB.Menu mnuEffectTop 
      Caption         =   "Effec&t"
      Begin VB.Menu mnuEffect 
         Caption         =   "Flip &horizontaly"
         Index           =   0
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Flip &verticaly"
         Index           =   1
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Rotate &left"
         Index           =   3
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "Rotate &right"
         Index           =   4
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "&Alpha"
         Index           =   6
         Begin VB.Menu mnuAlpha 
            Caption         =   "Drop &shadow"
            Index           =   0
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "&Fade down"
            Index           =   1
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "Scan&lines"
            Index           =   2
         End
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "&Color"
         Index           =   7
         Begin VB.Menu mnuColor 
            Caption         =   "&Colorize A"
            Index           =   0
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Greys"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "&Enhance"
         Index           =   8
         Begin VB.Menu mnuEnhance 
            Caption         =   "&Soften"
            Index           =   0
         End
         Begin VB.Menu mnuEnhance 
            Caption         =   "S&harpen"
            Index           =   1
         End
         Begin VB.Menu mnuEnhance 
            Caption         =   "&Despeckle"
            Index           =   2
         End
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuEffect 
         Caption         =   "&Fix mask..."
         Index           =   10
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "PSC &page..."
         Index           =   2
      End
   End
   Begin VB.Menu mnuFrameSelectionTop 
      Caption         =   "Frame selection"
      Visible         =   0   'False
      Begin VB.Menu mnuFrameSelection 
         Caption         =   "Merge"
         Index           =   0
      End
      Begin VB.Menu mnuFrameSelection 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFrameSelection 
         Caption         =   "Cut"
         Index           =   2
      End
      Begin VB.Menu mnuFrameSelection 
         Caption         =   "Copy"
         Index           =   3
      End
      Begin VB.Menu mnuFrameSelection 
         Caption         =   "Paste"
         Index           =   4
      End
      Begin VB.Menu mnuFrameSelection 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFrameSelection 
         Caption         =   "Cancel"
         Index           =   6
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Application:   iICO
' Author:        Carles P.V.
' Last revision: 2004.07.12
'================================================
'
' Thanks to:
'
'   - Vlad Vissoultchev
'   - Ron van Tilburg
'
' Special thanks to:
'
'   - Avery
'
'================================================
'
' History:
'
'   - 2004.06.19 (v1.0.0)
'     * First release of iICO.
'
'   - 2004.06.22 (v1.0.1)
'     * Fixed: JASC-palette loading:
'       Palette entries now forced to fill remaining
'       entries (256). Thanks to Robert Rayment.
'
'   - 2004.06.28 (v1.0.2)
'     * Minor BIG BUG fixed (cDIB): m_hDC was being
'       destroyed through DeleteObject.
'     * Minor BIG BUG fixed (ucColor): Pattern DIB
'       processed out of bounds. Thanks to Vlad
'       Vissoultchev again.
'       //This is what was causing W9x crashes.
'
'   - 2004.06.28 (v1.0.3)
'     * Added: Alpha effects now available on selection
'       frame too.
'
'   - 2004.07.06 (v1.0.4)
'     * Fixed: Undo/Redo toolbar/menu refreshing.
'
'   - 2004.07.12 (v1.0.5)
'     * Fixed: Extension incorrectly checked on save.
'================================================



Option Explicit

'========================================================================================
' Initialization/Termination
'========================================================================================

Private Sub Form_Initialize()
    
    '-- Initialize...
    Call mSettings.LoadSettings
    Call mMain.InitializeApp
            
    '-- Check command line
    If (Command$ <> vbNullString) Then
        '-- Load from command line
        Call pvLoadFromShell(Replace$(Command$, Chr$(34), vbNullString))
      Else
        '-- Create default [icon] format
        G_oICON.ResourceType = [rtIcon]
        Call pvInitialize(Blank:=True, Untitled:=True)
    End If
    Call mUndo.SetListSaved(Saved:=True)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    '-- Icon has been changed [?]
    Call pvSaveBeforeContinuing(Cancel)
    If (Cancel = 0) Then
        Call Unload(Me)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '-- End...
    Call mUndo.CleanListUndoHistory
    Call mSettings.SaveSettings
    Call mMain.TerminateApp
End Sub

'//

Private Sub Form_Resize()
    
    '-- Resize/locate controls
    Call mMain.fMain_Resize
End Sub

'========================================================================================
' Menus
'========================================================================================

Private Sub mnuFile_Click(Index As Integer)

  Dim Filename As String
  Dim Cancel   As Integer
  
    Select Case Index
    
        Case 0 '-- New...
      
            With fNew
                '-- Show dialog
                Call .Show(vbModal, Me)
                If (Not .Cancel) Then
                    '-- Icon has been changed [?]
                    Call pvSaveBeforeContinuing(Cancel)
                    If (Cancel = 0) Then
                        '-- New resource...
                        G_oICON.ResourceType = .ResourceType
                        Call pvInitialize(Blank:=True, Untitled:=True)
                        Call mUndo.SetListSaved(Saved:=True)
                    End If
                End If
            End With
      
        Case 1 '-- Load...
            
            Filename = mDialogFile.GetFileName(Me.hWnd, G_OpenSavePath, "Icon|*.ico|Cursor|*.cur", G_oICON.ResourceType, "Load")
        
            If (Len(Filename)) Then
                '-- Icon has been changed [?]
                Call pvSaveBeforeContinuing(Cancel)
                If (Cancel = 0) Then
                    '-- Store filename
                    G_OpenSavePath = Filename
                    '-- Load...
                    If (G_oICON.LoadFromFile(G_OpenSavePath)) Then
                        Call pvInitialize(Blank:=False, Untitled:=False)
                        Call mUndo.SetListSaved(Saved:=True)
                      Else
                        Call MsgBox("Unexpected error loading resource.", vbExclamation)
                        Call pvInitialize(Blank:=True, Untitled:=True)
                    End If
                End If
             End If
        
        Case 3 '-- Save
        
            If (Len(G_OpenSavePath)) Then
                If (Not G_oICON.SaveToFile(G_OpenSavePath)) Then
                    Call MsgBox("Unexpected error saving resource.", vbExclamation)
                  Else
                    Call mUndo.SetListSaved(Saved:=True)
                End If
              Else
                Call mnuFile_Click(4)
            End If
      
        Case 4 '-- Save as...
            
            If (ucIconInfo.TextFile = "[Untitled]") Then
                G_OpenSavePath = Left$(G_OpenSavePath, InStrRev(Filename, "\"))
            End If
            
            If (G_oICON.ResourceType = [rtIcon]) Then
                Filename = mDialogFile.GetFileName(Me.hWnd, G_OpenSavePath, "Icon|*.ico", , "Save", False)
              Else
                Filename = mDialogFile.GetFileName(Me.hWnd, G_OpenSavePath, "Cursor|*.cur", , "Save", False)
            End If
            
            If (Len(Filename)) Then
                '-- Correct extension/Store filename
                Call pvCheckExtension(Filename)
                G_OpenSavePath = Filename
                '-- Update file info
                ucIconInfo.TextFile = Filename
                Call ucIconInfo.Refresh
                '-- Save...
                If (Not G_oICON.SaveToFile(G_OpenSavePath)) Then
                    Call MsgBox("Unexpected error saving resource.", vbExclamation)
                  Else
                    Call mUndo.SetListSaved(Saved:=True)
               End If
            End If
            
        Case 6 '-- Import...
            
            '-- Icon has been changed [?]
            Call pvSaveBeforeContinuing(Cancel)
            If (Cancel = 0) Then
            
                Filename = mDialogFile.GetFileName(Me.hWnd, G_ImportPath, "All supported images|*.png;*.gif;*.wmf;*.emf;*.bmp;*.tif;*.tiff;*.jpg|PNG images|*.png|GIF images|*.gif|WMF/EMF images|*.wmf;*.emf|BMP images|*.bmp|TIFF images|*.tif;*.tiff|JPG images|*.jpg", 1, "Import")
            
                If (Len(Filename)) Then
                    '-- Store filename
                    G_ImportPath = Filename
                    '-- Import...
                    With fImport
                        If (.Import(Filename)) Then
                            Call .Show(vbModal, Me)
                            If (.Cancel = False) Then
                                Call pvInitialize(Blank:=False, Untitled:=True)
                                Call mUndo.SetListSaved(Saved:=False)
                            End If
                          Else
                            Call MsgBox("Unexpected error importing image.", vbExclamation)
                        End If
                    End With
                End If
            End If
            
        Case 7 '-- Explorer... To do
        
        Case 9 '-- Exit
            
            Call Unload(Me)
    End Select
End Sub

Private Sub mnuEdit_Click(Index As Integer)

    Select Case Index
    
        Case 0 '-- Undo
            Call mUndo.Undo(ImageIdx:=G_ImageIdx)
            Call pvUpdateItemInfo
            ucIconList.ListIndex = G_ImageIdx
        
        Case 1 '-- Redo
            Call mUndo.Redo(ImageIdx:=G_ImageIdx)
            Call pvUpdateItemInfo
            ucIconList.ListIndex = G_ImageIdx
            
        Case 3 '-- Cut
            Call ucIconCanvas.Cut
            If (Clipboard.GetFormat(vbCFDIB)) Then
                Call ucToolbarMain.EnableButton(7, True)
            End If
            
        Case 4 '-- Copy
            Call ucIconCanvas.Copy
            If (Clipboard.GetFormat(vbCFDIB)) Then
                Call ucToolbarMain.EnableButton(7, True)
            End If
        
        Case 5 '-- Paste
            If (Clipboard.GetFormat(vbCFDIB)) Then
                Call ucIconCanvas.Paste
                Call ucToolbarDrawTools.CheckButton(1, True)
                Call pvUpdateToolOptions
            End If
        
        Case 6 '-- Clear
            Call ucIconCanvas.ClearSelection
        
        Case 8 '-- Select all
            Call ucToolbarDrawTools.CheckButton(1, True)
            Call ucIconCanvas.SelectAll
            Call pvUpdateToolOptions
    End Select
End Sub

Private Sub mnuPaletteTop_Click()
    
  Dim bEnablePalette As Boolean
  Dim bEnableOptimal As Boolean
  Dim x As Long, W As Long
  Dim y As Long, H As Long
    
    Call ucIconCanvas.GetSelectionInfo(x, y, W, H)
    bEnablePalette = (G_oICON.BPP(G_ImageIdx) > [016_Colors] And W = 0 And H = 0)
    bEnableOptimal = (G_oICON.BPP(G_ImageIdx) > [256_Colors] And W = 0 And H = 0)
    
    mnuPalette(0).Enabled = bEnablePalette
    mnuPalette(1).Enabled = bEnablePalette
    mnuPalette(3).Enabled = bEnablePalette
    mnuPalettePredefined(5).Enabled = bEnableOptimal
End Sub

Private Sub mnuPalette_Click(Index As Integer)

  Dim Filename As String
  Dim oPal     As New cPalette
  
    Select Case Index
    
        Case 0 '-- Load...
        
            Filename = mDialogFile.GetFileName(Me.hWnd, G_PalettePath, "JASC palette|*.pal", , "Load")
        
            If (Len(Filename)) Then
                '-- Store filename
                G_PalettePath = Filename
                '-- Load
                If (oPal.LoadFromJASCFile(Filename)) Then
                    oPal.Entries = 256
                    Call oPal.BuildLogicalPalette
                    Call pvSetPalette(oPal)
                  Else
                    Call MsgBox("Unexpected error loading palette.", vbExclamation)
                End If
             End If
             
        Case 1 '-- Save...
                
            Filename = mDialogFile.GetFileName(Me.hWnd, , "JASC palette|*.pal", , "Save", False)
            
            If (Len(Filename)) Then
                
                '-- Build temp. palette for saving
                Call oPal.Initialize(256)
                Call CopyMemory(ByVal oPal.lpPalette, G_ColorInfo(G_ImageIdx).Palette(0), 1024)
                '-- Save
                If (oPal.SaveToJASCFile(Filename)) Then
                  Else
                    Call MsgBox("Unexpected error saving palette.", vbExclamation)
                End If
            End If
    End Select
End Sub

Private Sub mnuPalettePredefined_Click(Index As Integer)
   
  Dim oPal As New cPalette
  Dim oDIB As New cDIB
           
    '-- Build palette
    Select Case Index
        
        Case 0 '-- Web safe
            Call oPal.CreateWebsafe
            
        Case 1 '-- Grey scale
            Call oPal.CreateGreyScale([256_pgColors])
        
        Case 2 '-- Spectrum
            Call oPal.CreateSpectrum
        
        Case 3 '-- XP basic
            Call oPal.CreateWindowXPBasic: oPal.Entries = 256
        
        Case 5 '-- Optimal
            With G_oICON
                Call oDIB.Create(.Width(G_ImageIdx), .Height(G_ImageIdx), [32_bpp])
                Call oDIB.LoadBlt(.oXORDIB(G_ImageIdx).hDC)
            End With
            Call oPal.CreateOptimal(oDIB, 255, 8): oPal.Entries = 256
            Call oPal.SortPalette
            Call oPal.BuildLogicalPalette
    End Select
    
    '-- Set new icon/work palette ([Dither])
    Call pvSetPalette(oPal)
End Sub

Private Sub mnuCanvas_Click(Index As Integer)
  
  Dim lClr As Long
    
    Select Case Index
        
        Case 0 '-- Screen color...
        
            '-- Show color dialog
            lClr = mDialogColor.SelectColor(Me.hWnd, ucColorScreen.Color, True)
            
            '-- We have choosen a color
            If (lClr <> -1) Then
                
                '-- Update global Color Screen var.
                G_ColorScreen = lClr
                
                '-- Update color picker
                ucColorScreen.Color = G_ColorScreen
                If (ucColorA.IsColorScreen) Then ucColorA.Color = G_ColorScreen
                If (ucColorB.IsColorScreen) Then ucColorB.Color = G_ColorScreen
                
                '-- Update icon list
                ucIconList.ColorScreen = G_ColorScreen
                Call ucIconList.RefreshItem(G_ImageIdx)
                
                '-- Update canvas
                ucIconCanvas.CanvasEx.ColorScreen = G_ColorScreen
                Call ucIconCanvas.Refresh
                
                '-- Update preview window
                picPreview.BackColor = G_ColorScreen
                Call picPreview_Paint
            End If
        
        Case 1 '-- Pixel grid
        
            mnuCanvas(1).Checked = Not mnuCanvas(1).Checked
            
            '-- Update canvas
            ucIconCanvas.ShowGrid = mnuCanvas(1).Checked
            Call ucIconCanvas.Refresh
                        
        Case 3 '-- Antialias
        
            mnuCanvas(3).Checked = Not mnuCanvas(3).Checked
            
            '-- Update canvas prop.
            ucIconCanvas.CanvasEx.Antialias = mnuCanvas(3).Checked
            '-- Update canvas
            If (ucIconCanvas.Tool = [icText]) Then
                Call ucIconCanvas.Refresh
            End If
    End Select
End Sub

Private Sub mnuEffectTop_Click()
 
  Dim bEnableFlipRotation As Boolean
  Dim bEnableAlphaEffects As Boolean
      
    '-- Enable/Disable effects
    bEnableFlipRotation = (ucIconCanvas.Tool <> [icText])
    bEnableAlphaEffects = (ucIconCanvas.Tool <> [icText] And G_oICON.BPP(G_ImageIdx) = [ARGB_Color])
    
    mnuEffect(0).Enabled = bEnableFlipRotation
    mnuEffect(1).Enabled = bEnableFlipRotation
    mnuEffect(3).Enabled = bEnableFlipRotation
    mnuEffect(4).Enabled = bEnableFlipRotation
    mnuEffect(6).Enabled = bEnableAlphaEffects
    mnuEffect(7).Enabled = bEnableAlphaEffects
    mnuEffect(8).Enabled = bEnableAlphaEffects
    mnuEffect(10).Enabled = bEnableAlphaEffects
    
    '-- Disable Colorize if A = color Screen
    mnuColor(0).Enabled = Not ucColorA.IsColorScreen
End Sub

Private Sub mnuEffect_Click(Index As Integer)
  
    Select Case Index

        Case 0  '-- Flip horizontaly
            Call ucIconCanvas.FlipHorizontaly

        Case 1  '-- Flip verticaly
            Call ucIconCanvas.FlipVerticaly
            
        Case 3  '-- Rotate left
            Call picPreview.Cls
            Call ucIconCanvas.RotateLeft

        Case 4  '-- Rotate right
            Call picPreview.Cls
            Call ucIconCanvas.RotateRight
            
        Case 10 ' Fix mask...
            With fMaskFix
                Call .FixMask
                Call .Show(vbModal, Me)
                If (.Cancel = False) Then
                    Call mUndo.SaveUndo(ImageIdx:=G_ImageIdx)
                End If
            End With
    End Select
End Sub

Private Sub mnuColor_Click(Index As Integer)
  
  Dim x As Long, W As Long
  Dim y As Long, H As Long
    
    Call ucIconCanvas.GetSelectionInfo(x, y, W, H)
    
    If (ucIconCanvas.Tool = [icSelectionFrame] And W > 0 And H > 0) Then
        
        '-- Apply on selection
        Select Case Index
            Case 0 '-- Colorize
                Call mARGBFilter.Colorize(ucIconCanvas.oFrameSelectionXOR, ucColorA.Color)
            Case 1 '-- Greys
                Call mARGBFilter.Greys(ucIconCanvas.oFrameSelectionXOR)
        End Select
        Call ucIconCanvas.Refresh
    
      Else
        '-- Apply on icon
        Select Case Index
            Case 0 '-- Colorize
                Call mARGBFilter.Colorize(G_oICON.oXORDIB(G_ImageIdx), ucColorA.Color)
            Case 1 '-- Greys
                Call mARGBFilter.Greys(G_oICON.oXORDIB(G_ImageIdx))
        End Select
        Call ucIconCanvas.Initialize
        Call ucIconCanvas_IconChange
    End If
End Sub

Private Sub mnuEnhance_Click(Index As Integer)
  
  Dim x As Long, W As Long
  Dim y As Long, H As Long
    
    Call ucIconCanvas.GetSelectionInfo(x, y, W, H)
    
    If (ucIconCanvas.Tool = [icSelectionFrame] And W > 0 And H > 0) Then
      
        '-- Apply on selection
        Select Case Index
            Case 0 '-- Soften
                Call mARGBFilter.Soften(ucIconCanvas.oFrameSelectionXOR)
            Case 1 '-- Sharpen
                Call mARGBFilter.Sharpen(ucIconCanvas.oFrameSelectionXOR)
            Case 2 '-- Despeckle
                Call mARGBFilter.Despeckle(ucIconCanvas.oFrameSelectionXOR)
        End Select
        Call ucIconCanvas.Refresh
      
      Else
        '-- Apply on icon
        Select Case Index
            Case 0 '-- Soften
                Call mARGBFilter.Soften(G_oICON.oXORDIB(G_ImageIdx))
            Case 1 '-- Sharpen
                Call mARGBFilter.Sharpen(G_oICON.oXORDIB(G_ImageIdx))
            Case 2 '-- Despeckle
                Call mARGBFilter.Despeckle(G_oICON.oXORDIB(G_ImageIdx))
        End Select
        Call ucIconCanvas.Initialize
        Call ucIconCanvas_IconChange
    End If
End Sub

Private Sub mnuAlpha_Click(Index As Integer)
  
  Dim x As Long, W As Long
  Dim y As Long, H As Long
    
    Call ucIconCanvas.GetSelectionInfo(x, y, W, H)
    
    If (ucIconCanvas.Tool = [icSelectionFrame] And W > 0 And H > 0) Then
    
        '-- Apply on selection
        Select Case Index
            Case 0 '-- Drop shadow
                Call mARGBFilter.DropShadow(ucIconCanvas.oFrameSelectionXOR)
            Case 1 '-- Fade down alpha
                Call mARGBFilter.FadeDownAlpha(ucIconCanvas.oFrameSelectionXOR)
            Case 2 '-- Alpha scanlines
                Call mARGBFilter.AlphaScanlines(ucIconCanvas.oFrameSelectionXOR)
        End Select
        Call ucIconCanvas.Refresh
      Else
        '-- Apply on icon
        Select Case Index
            Case 0 '-- Drop shadow
                Call mARGBFilter.DropShadow(G_oICON.oXORDIB(G_ImageIdx))
            Case 1 '-- Fade down alpha
                Call mARGBFilter.FadeDownAlpha(G_oICON.oXORDIB(G_ImageIdx))
            Case 2 '-- Alpha scanlines
                Call mARGBFilter.AlphaScanlines(G_oICON.oXORDIB(G_ImageIdx))
        End Select
        Call ucIconCanvas.Initialize
        Call ucIconCanvas_IconChange
    End If
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    
    Select Case Index
    
        Case 0 '-- Quick and short About
            Call MsgBox("iICO " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
                        "for Windows NT/2000/XP" & Space$(10) & vbCrLf & vbCrLf & _
                        "Carles P.V. - ©2004", vbInformation)
                        
        Case 2 '-- PSC page...
        
            If (MsgBox("You will be navigated to the PSC page of iICO.", vbYesNo + vbInformation) = vbYes) Then
                Call mMisc.Navigate(Me.hWnd, "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=54524")
            End If
    End Select
End Sub

'========================================================================================
' Quick keys
'========================================================================================

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        Case vbKeyEscape
            
            '-- 'Reset'
            If (ucIconCanvas.Tool = [icSelectionFrame] Or _
                ucIconCanvas.Tool = [icText]) Then
                ucIconCanvas.Tool = ucIconCanvas.Tool
                Call ucIconCanvas.Initialize
            End If
            
        Case vbKeyDelete
            
            '-- Clear selection
            Call mnuEdit_Click(7)
    End Select
End Sub

'========================================================================================
' Main toolbar
'========================================================================================

Private Sub ucToolbarMain_ButtonClick(Index As Integer, x1 As Long, y1 As Long, x2 As Long, y2 As Long)

    Select Case Index
    
        Case 1  '-- New...
            Call mnuFile_Click(0)
        Case 2  '-- Open...
            Call mnuFile_Click(1)
        Case 3  '-- Save as...
            Call mnuFile_Click(4)
        Case 4  '-- Undo
            Call mnuEdit_Click(0)
        Case 5  '-- Redo
            Call mnuEdit_Click(1)
        Case 6  '-- Cut
            Call mnuEdit_Click(3)
        Case 7  '-- Copy
            Call mnuEdit_Click(4)
        Case 8  '-- Paste
            Call mnuEdit_Click(5)
        Case 9  '-- Screen capture
            Me.Enabled = False
            Call fCapture.Show(vbModeless, Me)
    End Select
End Sub

'========================================================================================
' Add/Remove format
'========================================================================================

Private Sub ucToolbarFormat_ButtonClick(Index As Integer, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    
  Dim nIndS As Integer
  
    Select Case Index
    
        Case 1 '-- Add format
        
            With fNewFormat
            
                '-- Set dialog mode / Show dialog
                Call .SetDialogMode(G_oICON.ResourceType)
                Call .Show(vbModal, Me)
                
                If (.Cancel = False) Then
                    
                    nIndS = G_oICON.AddFormat(.NewWidth, .NewHeight, .NewBPP)
                    If (nIndS > -1) Then
                        '-- Format added
                        Call pvInsertColorInfo(nIndS)
                        Call pvInitializeColorInfo(nIndS)
                        '-- Insert format Undo history / Save Undo
                        Call mUndo.InsertUndoItemHistory(ImageIdx:=nIndS)
                        '-- Refresh
                        Call pvInitializeIconList(SelectItem:=nIndS, InitializeColorInfo:=False)
                        Call pvUpdateUndoTools
                        Call pvUpdateFileInfo
                      Else
                        '-- Format already exists...
                        Call MsgBox("This format already exists. Choose another format.", vbInformation)
                    End If
                End If
            End With
            
        Case 2 '-- Remove format
        
            If (G_oICON.Count > 1) Then
        
                '-- Confirm
                If (MsgBox("Do you want to remove the <" & ucIconList.ItemText(G_ImageIdx) & "> format from the current icon ?" & vbCrLf & vbCrLf & _
                           "This action can not be undone.", _
                           vbYesNo Or vbExclamation _
                           ) = vbYes) Then
                    
                    If (G_oICON.RemoveFormat(G_ImageIdx)) Then
                        '-- Format removed
                        Call pvRemoveColorInfo(G_ImageIdx)
                        '-- Remove format Undo history
                        Call mUndo.RemoveUndoItemHistory(G_ImageIdx)
                        '-- Refresh
                        Call pvInitializeIconList(SelectItem:=IIf(G_ImageIdx < G_oICON.Count, G_ImageIdx, G_oICON.Count - 1), InitializeColorInfo:=False)
                        Call pvUpdateUndoTools
                        Call pvUpdateFileInfo
                   End If
                End If
                
              Else
                Call MsgBox("Cannot remove the only image in an icon.", vbInformation)
            End If
    End Select
End Sub

'========================================================================================
' Icon list control
'========================================================================================

Private Sub ucIconList_Click()
        
    '-- Store current icon image index
    G_ImageIdx = ucIconList.ListIndex
    
    '-- Update
    Call pvUpdateUndoTools
    Call pvUpdateColorPickers
    Call pvUpdatePalettePicker
    Call pvCheckPaletteAndInitializeColors
    Call frUpdateAlphaTools
    
    '-- Hotspot coords. [?]
    lblHotSpotV.Caption = G_oICON.HotSpotX(G_ImageIdx) & "," & G_oICON.HotSpotY(G_ImageIdx)
    
    '-- Clear preview window / Initialize canvas
    Call picPreview.Cls
    Call ucIconCanvas.SetIconImageIndex(G_ImageIdx)
    Call ucIconCanvas.Initialize
End Sub

'========================================================================================
' Selectors / palette
'========================================================================================

Private Sub ucColorScreen_Click(ByVal Button As Integer, ByVal Shift As Integer)
      
    Select Case Button
        
        Case vbLeftButton
            ucColorA.IsColorScreen = True
            ucColorA.Color = G_ColorScreen
            ucIconCanvas.CanvasEx.IsAScreen = True
            ucIconCanvas.CanvasEx.SwapColors = False
        
        Case vbRightButton
            ucColorB.IsColorScreen = True
            ucColorB.Color = G_ColorScreen
            ucIconCanvas.CanvasEx.IsBScreen = True
            ucIconCanvas.CanvasEx.SwapColors = True
    End Select
    
    If (ucIconCanvas.Tool = [icText]) Then
        Call ucIconCanvas.Refresh
    End If
End Sub

Private Sub ucColorA_Click(ByVal Button As Integer, ByVal Shift As Integer)
    
    ucIconCanvas.CanvasEx.SwapColors = False
End Sub

Private Sub ucColorB_Click(ByVal Button As Integer, ByVal Shift As Integer)
    
    ucIconCanvas.CanvasEx.SwapColors = True
End Sub

'========================================================================================
' Alpha picker
'========================================================================================

Private Sub ucAlphaPicker_Change(ByVal Alpha As Byte)
    
    ucAlphaPicker.ToolTipText = "Alpha = " & Alpha
    ucColorA.Alpha = Alpha: Call ucColorA.Refresh
    ucColorB.Alpha = Alpha: Call ucColorB.Refresh
    ucIconCanvas.CanvasEx.Alpha = Alpha: Call ucIconCanvas.Refresh
End Sub

'========================================================================================
' Palette picker
'========================================================================================

Private Sub ucPalettePicker_ColorOver(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal Index As Byte)

    ucIconInfo.TextInfo3 = "RGB: " & R & "," & G & "," & B & " Index: " & Index
    Call ucIconInfo.Refresh
End Sub

Private Sub ucPalettePicker_ColorASelected(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal Index As Byte)
    
    '-- Color info
    G_ColorInfo(G_ImageIdx).ColorIdxA = Index
    
    '-- Canvas
    ucColorA.IsColorScreen = False
    ucColorA.Color = RGB(R, G, B)
    ucIconCanvas.CanvasEx.IsAScreen = False
    ucIconCanvas.CanvasEx.ColorLngA = RGB(R, G, B)
    ucIconCanvas.CanvasEx.ColorIdxA = Index
    ucIconCanvas.CanvasEx.SwapColors = False
    
    If (ucIconCanvas.Tool = [icText]) Then
        Call ucIconCanvas.Refresh
    End If
End Sub

Private Sub ucPalettePicker_ColorBSelected(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal Index As Byte)
    
    '-- Color info
    G_ColorInfo(G_ImageIdx).ColorIdxB = Index
    
    '-- Canvas
    ucColorB.IsColorScreen = False
    ucColorB.Color = RGB(R, G, B)
    ucIconCanvas.CanvasEx.IsBScreen = False
    ucIconCanvas.CanvasEx.ColorLngB = RGB(R, G, B)
    ucIconCanvas.CanvasEx.ColorIdxB = Index
    ucIconCanvas.CanvasEx.SwapColors = True
End Sub

Private Sub ucPalettePicker_ColorDblClick(Button As Integer, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal Index As Byte)
    
  Dim lClr   As Long
  Dim lIdx   As Long
  
    If (G_oICON.BPP(G_ImageIdx) >= [256_Colors]) Then
        
        If (RGB(R, G, B) = vbBlack And ucIconCanvas.IsPaletteBlackEntryAvailable(Index) = False) Then
            
            '-- Especial case:
            '   Black entry is being used for masking
            Call MsgBox("This entry is being used for masking or" & vbCrLf & _
                        "is the last black entry in the color palette." & vbCrLf & vbCrLf & _
                        "Choose another entry.", vbExclamation)
        
          Else
        
            '-- Show Color dialog
            lClr = mDialogColor.SelectColor(Me.hWnd, RGB(R, G, B), True)
            
            If (lClr <> -1) Then
            
                With G_ColorInfo(G_ImageIdx)
                    
                    '-- Change entry
                    .Palette(4 * Index + 2) = (lClr And &HFF&)
                    .Palette(4 * Index + 1) = (lClr And &HFF00&) \ 256
                    .Palette(4 * Index + 0) = (lClr And &HFF0000) \ 65536
                    '-- Change index
                    Select Case Button
                        Case vbLeftButton:  .ColorIdxA = Index
                        Case vbRightButton: .ColorIdxB = Index
                    End Select
                    
                    '-- Update icon palette [?]
                    If (G_oICON.BPP(G_ImageIdx) = [256_Colors]) Then
                        Call G_oICON.oXORDIB(G_ImageIdx).SetPalette(.Palette())
                        Call ucIconList.RefreshItem(G_ImageIdx)
                        '-- Initialize canvas (buffers)
                        Call ucIconCanvas.Initialize
                    End If
                    
                    '-- Update color pickers
                    Call pvUpdateColorPickers
                    
                    '-- Update palette picker
                    Call ucPalettePicker.SetPalette(.Palette())
                    Call ucPalettePicker.Refresh
                End With
                
                '-- Save Undo
                Call mUndo.SaveUndo(ImageIdx:=G_ImageIdx)
                Call pvUpdateUndoTools
            End If
        End If
    End If
End Sub

Private Sub ucPalettePicker_MouseOut()
    
    '-- Clear color info
    ucIconInfo.TextInfo3 = vbNullString
    Call ucIconInfo.Refresh
End Sub

'========================================================================================
' Drawing tools toolbar + tool options + menus
'========================================================================================

Private Sub ucToolbarDrawTools_ButtonClick(Index As Integer, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
                                      
    With ucIconCanvas
        
        '-- Select tool / Initialize canvas
        If (.Tool <> Index) Then
            .Tool = Index
            If (.Tool <> [icColorSelector]) Then
                Call .Initialize
              Else
                Call .Refresh
            End If
        End If
    
        '-- Store
        If (.Tool <> [icColorSelector] And .Tool <> [icHotSpot]) Then
            G_LastTool = .Tool
        End If
    End With
    
    '-- Tool options
    Call pvUpdateToolOptions
End Sub

Private Sub chkImportColorBAsScreen_Click()
    ucIconCanvas.CanvasEx.MaskImport = -chkImportColorBAsScreen
End Sub

Private Sub chkExportScreenAsColorB_Click()
    ucIconCanvas.CanvasEx.MaskExport = -chkExportScreenAsColorB
End Sub

Private Sub chkImport32bppOpaque_Click()
    ucIconCanvas.CanvasEx.Import32bppOpaque = -chkImport32bppOpaque
End Sub

Private Sub ucToolbarPixelMask_ButtonClick(Index As Integer, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    ucIconCanvas.CanvasEx.PixelMask = Index - 1
End Sub

Private Sub cbStraightLineWidth_Click()
    ucIconCanvas.CanvasEx.StraightLineWidth = 2 * cbStraightLineWidth.ListIndex + 1
    Call ucIconCanvas.Refresh(DrawToolPointer:=True)
End Sub

Private Sub cbBrushLineWidth_Click()
    ucIconCanvas.CanvasEx.BrushLineWidth = 2 * cbBrushLineWidth.ListIndex + 1
    Call ucIconCanvas.Refresh(DrawToolPointer:=True)
End Sub

Private Sub cbShape_Click()
    ucIconCanvas.CanvasEx.Shape = cbShape.ListIndex
    Call ucIconCanvas.Refresh(DrawToolPointer:=True)
End Sub

Private Sub cbShapeLineWidth_Click()
    ucIconCanvas.CanvasEx.ShapeLineWidth = 2 * cbShapeLineWidth.ListIndex + 1
    Call ucIconCanvas.Refresh(DrawToolPointer:=True)
End Sub

Private Sub ucARGBTolerance_Change(Index As Integer)
    
    With ucIconCanvas.CanvasEx
        Select Case Index
            Case 0: .FillToleranceA = ucARGBTolerance(0).Value
            Case 1: .FillToleranceH = ucARGBTolerance(1).Value
            Case 2: .FillToleranceS = ucARGBTolerance(2).Value
            Case 3: .FillToleranceL = ucARGBTolerance(3).Value
        End Select
    End With
End Sub

Private Sub txtText_GotFocus()
    
    txtText.SelStart = 0
    txtText.SelLength = Len(txtText)
End Sub

Private Sub txtText_Change()
    
    ucIconCanvas.CanvasEx.Text = txtText.Text
    Call ucIconCanvas.Refresh(DrawToolPointer:=False)
End Sub

Private Sub cbFont_Click()
    
    ucIconCanvas.CanvasEx.Font.Name = cbFont.List(cbFont.ListIndex)
    ucIconCanvas.CanvasEx.Font.Size = cbFontSize.List(cbFontSize.ListIndex)
    ucIconCanvas.CanvasEx.Font.Bold = ucToolbarFont.IsButtonChecked(1)
    ucIconCanvas.CanvasEx.Font.Italic = ucToolbarFont.IsButtonChecked(2)
    Call ucIconCanvas.Refresh(DrawToolPointer:=False)
End Sub

Private Sub cbFontSize_Click()
    
    ucIconCanvas.CanvasEx.Font.Size = cbFontSize.List(cbFontSize.ListIndex)
    Call ucIconCanvas.Refresh(DrawToolPointer:=False)
End Sub

Private Sub ucToolbarFont_ButtonCheck(Index As Integer, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    
    ucIconCanvas.CanvasEx.Font.Bold = ucToolbarFont.IsButtonChecked(1)
    ucIconCanvas.CanvasEx.Font.Italic = ucToolbarFont.IsButtonChecked(2)
    Call ucIconCanvas.Refresh(DrawToolPointer:=False)
End Sub

'========================================================================================
' Icon canvas control
'========================================================================================

Private Sub ucIconCanvas_MouseDown(Button As Integer, Shift As Integer, ByVal x As Long, ByVal y As Long, ByVal InSelection As Boolean)
    
    Select Case ucIconCanvas.Tool

        Case [icSelectionFrame], [icText]
                
            If (Button = vbRightButton And InSelection) Then
                
                '-- Enable available items
                With ucIconCanvas
                    mnuFrameSelection(2).Visible = (.Tool = [icSelectionFrame])
                    mnuFrameSelection(3).Visible = (.Tool = [icSelectionFrame])
                    mnuFrameSelection(4).Visible = (.Tool = [icSelectionFrame])
                    mnuFrameSelection(5).Visible = (.Tool = [icSelectionFrame])
                End With
                '-- Show context menu
                Call Me.PopupMenu(mnuFrameSelectionTop, , , , mnuFrameSelection(0))
             End If
                  
         Case [icHotSpot]
            
            '-- New hotspot coords.
            G_oICON.HotSpotX(G_ImageIdx) = x
            G_oICON.HotSpotY(G_ImageIdx) = y
            lblHotSpotV.Caption = G_oICON.HotSpotX(G_ImageIdx) & "," & G_oICON.HotSpotY(G_ImageIdx)
    End Select
End Sub

Private Sub ucIconCanvas_MouseMove(Button As Integer, Shift As Integer, ByVal x As Long, ByVal y As Long, ByVal InSelection As Boolean)
    
  Dim sInfo    As String
  
  Dim x1 As Long, x2 As Long, W As Long
  Dim y1 As Long, y2 As Long, H As Long
  
  Dim BPP      As dibBPPCts
  Dim IsScreen As Boolean
  Dim Index    As Byte
  Dim R        As Byte
  Dim G        As Byte
  Dim B        As Byte
  Dim A        As Byte
    
    '-- Mouse info:
    Call ucIconCanvas.GetMouseInfo(x1, y1, x2, y2)
    
    '-- Into image area [?]
    If (x2 >= 0 And y2 >= 0 And x2 < G_oICON.Width(G_ImageIdx) And y2 < G_oICON.Height(G_ImageIdx)) Then
        
        '-- Build info text...
        Select Case ucIconCanvas.Tool
            
            Case [icSelectionFrame]
                
                If (InSelection) Then
                    Call ucIconCanvas.GetSelectionInfo(x1, y1, W, H)
                    ucIconInfo.TextInfo2 = x1 & "," & y1 & " " & W & "x" & H
                  Else
                    Call ucIconCanvas.GetMouseInfo(x1, y1, x2, y2)
                    If (Button = vbLeftButton) Then
                        ucIconInfo.TextInfo2 = x1 & "," & y1 & " " & Abs(x2 - x1) + 1 & "x" & Abs(y2 - y1) + 1
                      Else
                        ucIconInfo.TextInfo2 = x2 & "," & y2
                    End If
                End If
                
            Case [icPencil], [icBrush], [icFloodFill]
            
                ucIconInfo.TextInfo2 = x & "," & y
            
            Case [icStraightLine]
                
                Call ucIconCanvas.GetMouseInfo(x1, y1, x2, y2)
                If (Button) Then
                    ucIconInfo.TextInfo2 = x1 & "," & y1 & " " & x2 & "," & y2
                  Else
                    ucIconInfo.TextInfo2 = x2 & "," & y2
                End If
                
            Case [icShape]
                
                Call ucIconCanvas.GetMouseInfo(x1, y1, x2, y2)
                If (Button) Then
                    ucIconInfo.TextInfo2 = x1 & "," & y1 & " " & Abs(x2 - x1) + 1 & "x" & Abs(y2 - y1) + 1
                  Else
                    ucIconInfo.TextInfo2 = x2 & "," & y2
                End If
            
            Case [icText]
                
                If (InSelection) Then
                    Call ucIconCanvas.GetSelectionInfo(x1, y1, W, H)
                    ucIconInfo.TextInfo2 = x1 & "," & y1 & " " & W & "x" & H
                  Else
                    ucIconInfo.TextInfo2 = vbNullString
                End If
           
            Case [icColorSelector]
        
                Call ucIconCanvas.GetPixelInfo(x, y, BPP, R, G, B, A, Index, IsScreen)
                
                If (IsScreen) Then
                    
                    sInfo = "Screen"
                    Call ucPalettePicker.SetCursor(0, 0)
                    
                  Else
                
                    sInfo = "RGB: " & R & "," & G & "," & B
                    
                    Select Case BPP
                        
                        Case [01_bpp], [04_bpp], [08_bpp]
                            sInfo = sInfo & " Entry: " & Index
                            sInfo = sInfo & IIf(IsScreen, "(Screen)", vbNullString)
                            Call ucPalettePicker.SetCursor(Index, True)
                       
                        Case [24_bpp]
                            sInfo = sInfo & IIf(IsScreen, " (Screen)", vbNullString)
                        
                        Case [32_bpp]
                            sInfo = sInfo & " Alpha: " & A
                    End Select
                End If
                
                ucIconInfo.TextInfo2 = x & "," & y
                ucIconInfo.TextInfo3 = sInfo
                
            Case [icHotSpot]
                
                ucIconInfo.TextInfo2 = x & "," & y
        End Select
        
      Else
        ucIconInfo.TextInfo2 = vbNullString
    End If
    
    '-- Refresh info text
    Call ucIconInfo.Refresh
End Sub

Private Sub ucIconCanvas_MouseUp(Button As Integer, Shift As Integer, ByVal x As Long, ByVal y As Long, ByVal InSelection As Boolean)

  Dim BPP      As dibBPPCts
  Dim IsScreen As Boolean
  Dim Index    As Byte
  Dim R        As Byte
  Dim G        As Byte
  Dim B        As Byte
  Dim A        As Byte
  
    Select Case ucIconCanvas.Tool
        
        Case [icColorSelector]
            
            '-- Pixel info:
            Call ucIconCanvas.GetPixelInfo(x, y, BPP, R, G, B, A, Index, IsScreen)
            
            '-- Screen layer [?]
            IsScreen = IsScreen Or (G_oICON.BPP(G_ImageIdx) = [ARGB_Color] And A = 0)
            
            '-- Build info text...
            Select Case Button
                
                Case vbLeftButton
                    
                    If (IsScreen) Then
                        ucColorA.Color = G_ColorScreen
                      Else
                        ucColorA.Color = RGB(R, G, B)
                        ucIconCanvas.CanvasEx.ColorLngA = RGB(R, G, B)
                        ucIconCanvas.CanvasEx.ColorIdxA = Index
                    End If
                    ucColorA.IsColorScreen = IsScreen
                    ucIconCanvas.CanvasEx.IsAScreen = IsScreen
                    ucIconCanvas.CanvasEx.SwapColors = False
                    
                    If (G_oICON.BPP(G_ImageIdx) = [ARGB_Color] And chkPickAlpha) Then
                        ucAlphaPicker.Value = A
                        Call ucPalettePicker.Refresh
                    End If
                    
                Case vbRightButton
                
                    If (IsScreen) Then
                        ucColorB.Color = G_ColorScreen
                      Else
                        ucColorB.Color = RGB(R, G, B)
                        ucIconCanvas.CanvasEx.ColorLngB = RGB(R, G, B)
                        ucIconCanvas.CanvasEx.ColorIdxB = Index
                    End If
                    ucColorB.IsColorScreen = IsScreen
                    ucIconCanvas.CanvasEx.IsBScreen = IsScreen
                    ucIconCanvas.CanvasEx.SwapColors = True
                    
                    If (G_oICON.BPP(G_ImageIdx) = [ARGB_Color] And chkPickAlpha) Then
                        ucAlphaPicker.Value = A
                        Call ucPalettePicker.Refresh
                    End If
            End Select
            
            Call ucPalettePicker.SetCursor(0, 0)
            
            ucIconInfo.TextInfo3 = vbNullString
            Call ucIconInfo.Refresh

            GoTo ChangeToPreviousTool
            
        Case [icHotSpot]
                
            GoTo ChangeToPreviousTool
    End Select
    Exit Sub

ChangeToPreviousTool:

    '-- Select previous tool
    ucIconCanvas.Tool = G_LastTool
    Call ucToolbarDrawTools.CheckButton(G_LastTool, True)
    '-- Tool options
    Call pvUpdateToolOptions
    
    '-- Refresh canvas
    Call ucIconCanvas.Refresh(DrawToolPointer:=False)
End Sub

Private Sub ucIconCanvas_MouseOut()
    
    '-- Hide palette entry cursor
    Call ucPalettePicker.SetCursor(0, 0)
    
    '-- Clear coords. and color info
    ucIconInfo.TextInfo2 = vbNullString
    ucIconInfo.TextInfo3 = vbNullString
    Call ucIconInfo.Refresh
End Sub

'-- Icon has changed (Edit):
Private Sub ucIconCanvas_IconChange()
    
    '-- Update item
    Call pvUpdateItemInfo
    Call ucIconList.RefreshItem(G_ImageIdx)
    
    '-- Save Undo
    Call mUndo.SaveUndo(ImageIdx:=G_ImageIdx)
    Call pvUpdateUndoTools
End Sub

'-- Canvas has changed:
Private Sub ucIconCanvas_CanvasChange()
    
    '-- Refresh preview view
    Call picPreview_Paint
End Sub

'-- Selection has changed:
Private Sub ucIconCanvas_SelectionChange()

  Dim x1 As Long, W As Long
  Dim y1 As Long, H As Long
    
    '-- Get selection coords.
    Call ucIconCanvas.GetSelectionInfo(x1, y1, W, H)
    
    '-- Update menu and toolbar
    mnuEdit(3).Enabled = (W > 0 And H > 0)
    mnuEdit(6).Enabled = (W > 0 And H > 0)
    Call ucToolbarMain.EnableButton(6, (W > 0 And H > 0))
End Sub

'-- Refresh preview:
Private Sub picPreview_Paint()

  Dim x As Long, y As Long
    
    If (Not G_oICON Is Nothing) Then
        x = 0.5 * (picPreview.ScaleWidth - G_oICON.Width(G_ImageIdx))
        y = 0.5 * (picPreview.ScaleHeight - G_oICON.Height(G_ImageIdx))
        Call ucIconCanvas.PaintCanvas(picPreview.hDC, x, y)
    End If
End Sub

Private Sub mnuFrameSelection_Click(Index As Integer)
    
    Select Case Index
        
        Case 0 '-- Merge selection (Frame/Text)
            Call ucIconCanvas.MergeSelection
        
        Case 2 '-- Cut
            Call mnuEdit_Click(3)
        
        Case 3 '-- Copy
            Call mnuEdit_Click(4)
        
        Case 4 '-- Paste
            Call mnuEdit_Click(5)
    End Select
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvLoadFromShell(ByVal Filename As String)
    
    '-- Load resource from Command line
    If (G_oICON.LoadFromFile(Filename)) Then
        G_OpenSavePath = Filename
        Call pvInitialize(Blank:=False, Untitled:=False)
      Else
        Call MsgBox("Unexpected error opening resource.", vbExclamation)
        Call pvInitialize(Blank:=True, Untitled:=True)
    End If
End Sub

Private Sub pvCheckExtension(Filename As String)
    
    '-- Check file extension
    If (LCase$(Right$(Filename, 4)) <> LCase$(Right$(mDialogFile.Extension, 4))) Then
        Filename = Filename & LCase$(Right$(mDialogFile.Extension, 4))
    End If
End Sub

'//

Private Sub pvInitialize(Optional ByVal Blank As Boolean = False, Optional ByVal Untitled As Boolean = True)
    
    '-- Clear and add default resource
    If (Blank) Then
        With G_oICON
            Call .Destroy
            Select Case .ResourceType
                Case [rtIcon]
                    Call .AddFormat(32, 32, [256_Colors])
                Case [rtCursor]
                    Call .AddFormat(32, 32, [002_Colors])
            End Select
        End With
    End If
    '-- Clean Undo history
    Call mUndo.CleanListUndoHistory
    
    '-- Enable/Disable <Set hot spot> tool
    Call ucToolbarDrawTools.EnableButton(9, G_oICON.ResourceType = [rtCursor])
    
    '-- Refresh list, file info, initialize Undo history
    Call pvInitializeIconList(SelectItem:=0, InitializeColorInfo:=True)
    Call pvUpdateFileInfo(Untitled:=Untitled)
    Call mUndo.InitializeListUndoHistory
End Sub

Private Sub pvInitializeIconList(Optional ByVal SelectItem As Integer = -1, _
                                 Optional ByVal InitializeColorInfo As Boolean = True)

  Dim TopItem As Integer
    
    '-- Hide list
    ucIconList.Visible = False
    
    '-- Store current top item index
    TopItem = ucIconList.TopIndex
    
    '-- Clear list and initialize format info array
    Call ucIconList.Clear
    If (InitializeColorInfo) Then
        ReDim G_ColorInfo(G_oICON.Count - 1)
    End If
    
    '-- Fill icon list
    For G_ImageIdx = 0 To G_oICON.Count - 1
        
        '-- Add item...
        Call ucIconList.AddItem(vbNullString)
        Call pvUpdateItemInfo
        
        '-- Initialize format info
        If (InitializeColorInfo) Then
            Call pvInitializeColorInfo(G_ImageIdx)
        End If
    Next G_ImageIdx
    
    '-- Select item [?]
    If (SelectItem > -1) Then
        With ucIconList
            .ListIndex = SelectItem
            Select Case True
                Case SelectItem > TopItem + .VisibleRows - 1 + (Not .PerfectRowPad)
                Case SelectItem < TopItem
                    .TopIndex = SelectItem
                Case Else
                    .TopIndex = TopItem
            End Select
        End With
    End If
    
    '-- Show (refresh) list
    ucIconList.Visible = True
End Sub

Private Sub pvUpdateItemInfo()
  
  Dim sItm As String
  
    '-- Item format
    sItm = G_oICON.Width(G_ImageIdx) & "x" & G_oICON.Height(G_ImageIdx) & Chr$(32)
    Select Case G_oICON.BPP(G_ImageIdx)
        Case [002_Colors]: sItm = sItm & "Monochrome"
        Case [016_Colors]: sItm = sItm & "16 colors"
        Case [256_Colors]: sItm = sItm & "256 colors"
        Case [True_Color]: sItm = sItm & "True color"
        Case [ARGB_Color]: sItm = sItm & "ARGB color"
    End Select
    ucIconList.ItemText(G_ImageIdx) = sItm
End Sub

'//

Private Sub pvInitializeColorInfo(ByVal ImageIdx As Integer)
    
  Dim oPal As New cPalette
  
    With G_ColorInfo(ImageIdx)
        
        '-- Palette:
        Select Case G_oICON.BPP(ImageIdx)
        
            Case [002_Colors], [016_Colors], [256_Colors]
                ReDim .Palette(4 * 2 ^ G_oICON.BPP(ImageIdx) - 1)
                Call G_oICON.oXORDIB(ImageIdx).GetPalette(.Palette())
        
            Case Else
                ReDim .Palette(4 * 2 ^ 8 - 1)
                Call oPal.CreateSpectrum
                Call CopyMemory(.Palette(0), ByVal oPal.lpPalette, 4 * 2 ^ 8)
        End Select
        
        '-- Reset Colors A/B (indexes to palette)
        .ColorIdxA = -1
        .ColorIdxB = -1
    End With
End Sub

Private Sub pvInsertColorInfo(ByVal ImageIdx As Integer)

  Dim lIdx As Long
  
    ReDim Preserve G_ColorInfo(UBound(G_ColorInfo()) + 1)
    For lIdx = UBound(G_ColorInfo()) - 1 To ImageIdx Step -1
        G_ColorInfo(lIdx + 1) = G_ColorInfo(lIdx)
    Next lIdx
End Sub

Private Sub pvRemoveColorInfo(ByVal ImageIdx As Integer)

  Dim lIdx As Long
    
    For lIdx = ImageIdx To UBound(G_ColorInfo()) - 1 Step 1
        G_ColorInfo(lIdx) = G_ColorInfo(lIdx + 1)
    Next lIdx
    ReDim Preserve G_ColorInfo(UBound(G_ColorInfo()) - 1)
End Sub

'//

Private Sub pvUpdatePalettePicker()
    
    '-- Load work palette and refresh
    Call ucPalettePicker.SetPalette(G_ColorInfo(G_ImageIdx).Palette())
    Call ucPalettePicker.Refresh
End Sub

Private Sub pvUpdateColorPickers()
  
    With G_ColorInfo(G_ImageIdx)
        
        '-- Update A [?]
        If (Not ucColorA.IsColorScreen And .ColorIdxA > -1) Then
            ucColorA.Color = RGB(.Palette(4 * .ColorIdxA + 2), .Palette(4 * .ColorIdxA + 1), .Palette(4 * .ColorIdxA + 0))
            ucIconCanvas.CanvasEx.ColorLngA = ucColorA.Color
        End If
        '-- Update B [?]
        If (Not ucColorB.IsColorScreen And .ColorIdxB > -1) Then
            ucColorB.Color = RGB(.Palette(4 * .ColorIdxB + 2), .Palette(4 * .ColorIdxB + 1), .Palette(4 * .ColorIdxB + 0))
            ucIconCanvas.CanvasEx.ColorLngB = ucColorB.Color
        End If
    End With
End Sub

Private Sub pvUpdateFileInfo(Optional ByVal Untitled As Boolean = False)
    
    '-- Show file name / image Count and total images size
    With ucIconInfo
        .TextFile = IIf(Untitled, "[Untitled]", G_OpenSavePath)
        .TextInfo1 = G_oICON.Count & " image/s: " & Format(G_oICON.ImagesSize, "#,# bytes")
        Call .Refresh
    End With
End Sub

'//

Private Sub pvCheckPaletteAndInitializeColors()
    
  Dim oPal      As cPalette
  Dim aIdxBlack As Byte, aIdxA As Byte
  Dim aIdxWhite As Byte, aIdxB As Byte
  
    '-- Create logical palette and get closest
    '   indexes to pure black and white colors
    
    With G_ColorInfo(G_ImageIdx)
        
        Set oPal = New cPalette
        Call oPal.Initialize((UBound(.Palette()) - LBound(.Palette()) + 1) \ 4)
        Call CopyMemory(ByVal oPal.lpPalette, .Palette(0), 4 * oPal.Entries)
        Call oPal.BuildLogicalPalette
        
        Call oPal.ClosestIndex(&H0, &H0, &H0, aIdxBlack)
        Call oPal.ClosestIndex(&HFF, &HFF, &HFF, aIdxWhite)
    End With
    
    '-- Check for pure black color (white -inverse- not supported here)
    
    If (G_oICON.BPP(G_ImageIdx) <= [256_Colors]) Then
        
        With oPal
            If (RGB(.rgbR(aIdxBlack), oPal.rgbG(aIdxBlack), oPal.rgbB(aIdxBlack)) <> vbBlack) Then
                Call MsgBox("The color palette does not include the color black." & vbCrLf & vbCrLf & _
                            "Use of 'screen' color will be unpredictable until a true black entry (0,0,0)" & vbCrLf & _
                            "has been defined in the color palette.", _
                             vbExclamation)
            End If
'           If (RGB(.rgbR(aIdxWhite), .rgbG(aIdxWhite), .rgbB(aIdxWhite)) <> vbWhite) Then
'               Call MsgBox("The color palette does not include the color white." & vbCrLf & vbCrLf & _
'                           "Use of 'inverse' color will be unpredictable until a true white entry (255,255,255)" & vbCrLf & _
'                           "has been defined in the color palette.", _
'                            vbExclamation)
'           End If
        End With
    End If
    
    '-- Assign colors (A/B)
    
    With G_ColorInfo(G_ImageIdx)
        aIdxA = IIf(.ColorIdxA > -1, .ColorIdxA, aIdxBlack)
        aIdxB = IIf(.ColorIdxB > -1, .ColorIdxB, aIdxWhite)
    End With
    With ucIconCanvas.CanvasEx
        If (.IsAScreen = False) Then
            .ColorIdxA = aIdxA
            .ColorLngA = RGB(oPal.rgbR(aIdxA), oPal.rgbG(aIdxA), oPal.rgbB(aIdxA))
            ucColorA.Color = .ColorLngA
        End If
        If (.IsBScreen = False) Then
            .ColorIdxB = aIdxB
            .ColorLngB = RGB(oPal.rgbR(aIdxB), oPal.rgbG(aIdxB), oPal.rgbB(aIdxB))
            ucColorB.Color = .ColorLngB
        End If
        .SwapColors = False
    End With
End Sub

Private Sub pvSetPalette(oPal As cPalette)

  Dim oDIBIn     As New cDIB
  Dim oDIBDIther As New cDIBDither
  Dim sRet       As String
  Dim bDither    As Boolean

    With G_oICON
        
        '-- Prepare source 32-bpp DIB
        Call oDIBIn.Create(.Width(G_ImageIdx), .Height(G_ImageIdx), [32_bpp])
        Call oDIBIn.LoadBlt(.oXORDIB(G_ImageIdx).hDC)
        
        '-- Dither [?]
        If (.BPP(G_ImageIdx) <= [256_Colors]) Then
            bDither = True
          Else
            sRet = MsgBox("Dither to palette ?", vbYesNo + vbInformation + vbDefaultButton2)
            bDither = (sRet = vbYes)
        End If
        If (bDither) Then
            If (oPal.IsGreyScale) Then
                Call oDIBDIther.DitherToGreyScale(oDIBIn, .oXORDIB(G_ImageIdx), Diffuse:=False)
              Else
                Call oDIBDIther.DitherToColorPalette(oPal, oDIBIn, .oXORDIB(G_ImageIdx), Diffuse:=True)
            End If
        End If
    End With
    
    With G_ColorInfo(G_ImageIdx)
    
        '-- Set work palette
        ReDim .Palette(4 * oPal.Entries - 1)
        Call CopyMemory(.Palette(0), ByVal oPal.lpPalette, 4 * oPal.Entries)
        Call pvCheckPaletteAndInitializeColors
        
        '-- Color pickers A/B
        Call pvUpdateColorPickers

        '-- Palette picker
        Call ucPalettePicker.SetPalette(.Palette())
        Call ucPalettePicker.Refresh
        
        '-- Icon palette [?]
        If (G_oICON.BPP(G_ImageIdx) = [256_Colors]) Then
            Call G_oICON.oXORDIB(G_ImageIdx).SetPalette(.Palette())
        End If
    End With

    '-- Update canvas [?]
    If (bDither) Then
        Call ucIconCanvas.Initialize
        Call ucIconCanvas_IconChange
    End If
End Sub

'//

Private Sub pvUpdateUndoTools()
    
    '-- Menu
    mnuEdit(0).Enabled = mUndo.IsItemUndoAvailable(G_ImageIdx)
    mnuEdit(1).Enabled = mUndo.IsItemRedoAvailable(G_ImageIdx)
    
    '-- Toolbar
    Call ucToolbarMain.EnableButton(4, mnuEdit(0).Enabled)
    Call ucToolbarMain.EnableButton(5, mnuEdit(1).Enabled)
End Sub

Private Sub pvUpdateToolOptions()

  Dim nFrm As Integer
  
    '-- Show tool options frame
    For nFrm = 1 To fraDrawTools.Count
        fraDrawTools(nFrm).Visible = False
    Next nFrm
    fraDrawTools(ucIconCanvas.Tool).Visible = True
    
    '-- Set focus
    On Error Resume Next
    Select Case ucIconCanvas.Tool
        Case 3: Call cbStraightLineWidth.SetFocus
        Case 4: Call cbBrushLineWidth.SetFocus
        Case 5: Call cbShape.SetFocus
        Case 6: Call ucARGBTolerance(0).SetFocus
        Case 7: Call txtText.SetFocus
    End Select
    On Error GoTo 0
End Sub

'//

Private Sub pvSaveBeforeContinuing(Cancel As Integer)
  
  Dim sRet As VbMsgBoxResult
    
    If (mUndo.IsListIrreversible Or mUndo.IsListSaved = False) Then
    
        sRet = MsgBox("Current" & IIf(G_oICON.ResourceType = [rtIcon], " icon ", " cursor ") & _
                      "has been changed." & vbCrLf & vbCrLf & _
                      "Save changes before continuing ?", _
                       vbYesNoCancel Or vbInformation)
                       
        Select Case sRet
            Case vbYes    '-- Save as...
                Call mnuFile_Click(4)
                Cancel = 0
            Case vbNo     '-- Don't save
                Cancel = 0
            Case vbCancel '-- Cancel
                Cancel = 1
        End Select
    End If
End Sub

'========================================================================================
' 'Friend'
'========================================================================================

Friend Sub frUpdateAlphaTools()

  Dim bIsARGB As Boolean
  
    '-- Is ARGB format [?]
    bIsARGB = (G_oICON.BPP(G_ImageIdx) = [ARGB_Color])
    
    '-- Enable/disable alpha tools:
    
    '-  Canvas menu
    mnuCanvas(3).Enabled = bIsARGB
    
    '-  Color A/B pickers
    ucColorA.Alpha = IIf(bIsARGB, ucAlphaPicker.Value, 255): Call ucColorA.Refresh
    ucColorB.Alpha = IIf(bIsARGB, ucAlphaPicker.Value, 255): Call ucColorB.Refresh
    
    '-  Alpha picker
    ucAlphaPicker.Enabled = bIsARGB
    
    '-- Import/Export
    chkImportColorBAsScreen.Enabled = (Clipboard.GetFormat(vbCFDIB) And Not ucIconCanvas.IsPrivateClipboardAvailable)
    chkImport32bppOpaque.Enabled = (bIsARGB And Clipboard.GetFormat(vbCFDIB) And Not ucIconCanvas.IsPrivateClipboardAvailable)
    chkExportScreenAsColorB.Enabled = Not bIsARGB
    
    '-  Alpha Flood-Fill options
    lblTolerance.Enabled = bIsARGB
    ucARGBTolerance(0).Enabled = bIsARGB
    ucARGBTolerance(1).Enabled = bIsARGB
    ucARGBTolerance(2).Enabled = bIsARGB
    ucARGBTolerance(3).Enabled = bIsARGB
    chkPickAlpha.Enabled = bIsARGB
End Sub
