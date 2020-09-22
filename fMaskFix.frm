VERSION 5.00
Begin VB.Form fMaskFix 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fix mask"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
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
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2595
      Width           =   1050
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2010
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2595
      Width           =   1050
   End
   Begin VB.Label lblAfter 
      Caption         =   "After"
      Height          =   225
      Left            =   2250
      TabIndex        =   1
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lblBefore 
      Caption         =   "Before"
      Height          =   225
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "fMaskFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- API:
Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const BDR_RAISEDINNER     As Long = &H4
Private Const BDR_SUNKENOUTER     As Long = &H2
Private Const BF_RECT             As Long = &HF
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT2, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long

'-- Private Variables:
Private m_rctBef    As RECT2
Private m_rctAft    As RECT2
Private m_oTile     As New cTile
Private m_oXORDIB   As New cDIB
Private m_oANDDIB   As New cDIB
Private m_oCanvasEx As New cIconCanvasEx
Private m_Cancel    As Boolean

'//

Public Sub FixMask()

    '-- Clone current ARGB format DIBs
    Call G_oICON.oXORDIB(G_ImageIdx).CloneTo(m_oXORDIB)
    Call G_oICON.oANDDIB(G_ImageIdx).CloneTo(m_oANDDIB)
    
    '-- Process...
    Call m_oCanvasEx.RemaskARGBIcon(m_oXORDIB, m_oANDDIB)
End Sub

Private Sub Form_Load()
    
    Set Me.Icon = Nothing

    Call mMisc.RemoveButtonBorderEnhance(Me.cmdOk)
    Call mMisc.RemoveButtonBorderEnhance(Me.cmdCancel)
    
    '-- Define Before and After rects.
    Call SetRect(m_rctBef, 0, 0, 130, 130): Call OffsetRect(m_rctBef, 10, 25)
    Call SetRect(m_rctAft, 0, 0, 130, 130): Call OffsetRect(m_rctAft, 150, 25)
    
    '-- Create 'transparent' layer
    Call m_oTile.CreatePatternFromStdPicture(LoadResPicture("PATTERN_8X8", vbResBitmap))
End Sub

Private Sub Form_Paint()
    
  Dim bfx As Long, bfy As Long
  Dim bfW As Long, bfH As Long
    
    '-- Draw edges and 'transparent' layer
    Call DrawEdge(hDC, m_rctBef, BDR_SUNKENOUTER, BF_RECT)
    Call DrawEdge(hDC, m_rctAft, BDR_SUNKENOUTER, BF_RECT)
    Call m_oTile.Tile(hDC, m_rctBef.x1 + 1, m_rctBef.y1 + 1, 128, 128)
    Call m_oTile.Tile(hDC, m_rctAft.x1 + 1, m_rctAft.y1 + 1, 128, 128)
    
    With G_oICON
                              
        bfW = .Width(G_ImageIdx)
        bfH = .Height(G_ImageIdx)
        
        '-- Fit image (128x128 max)
        Call m_oXORDIB.GetBestFitInfo(bfW, bfH, 128, 128, bfx, bfy, bfW, bfH)
        
        '-- Before
        Call m_oANDDIB.Stretch(hDC, bfx + m_rctAft.x1 + 1, bfy + m_rctAft.y1 + 1, bfW, bfH, , , , , vbSrcAnd)
        Call m_oXORDIB.Stretch(hDC, bfx + m_rctAft.x1 + 1, bfy + m_rctAft.y1 + 1, bfW, bfH, , , , , vbSrcPaint)
        '-- After
        Call .oANDDIB(G_ImageIdx).Stretch(hDC, bfx + m_rctBef.x1 + 1, bfy + m_rctBef.y1 + 1, bfW, bfH, , , , , vbSrcAnd)
        Call .oXORDIB(G_ImageIdx).Stretch(hDC, bfx + m_rctBef.x1 + 1, bfy + m_rctBef.y1 + 1, bfW, bfH, , , , , vbSrcPaint)
    End With
End Sub

'//

Private Sub cmdOk_Click()
    
    '-- Set new fixed DIBs
    Call m_oXORDIB.CloneTo(G_oICON.oXORDIB(G_ImageIdx))
    Call m_oANDDIB.CloneTo(G_oICON.oANDDIB(G_ImageIdx))
    Call m_oXORDIB.Destroy
    Call m_oANDDIB.Destroy
    
    m_Cancel = False
    Call Me.Hide
End Sub

Private Sub cmdCancel_Click()
    
    Call m_oXORDIB.Destroy
    Call m_oANDDIB.Destroy
    
    m_Cancel = True
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

