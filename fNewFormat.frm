VERSION 5.00
Begin VB.Form fNewFormat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New image format"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
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
   ScaleHeight     =   262
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSize 
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   135
      TabIndex        =   14
      Top             =   1605
      Width           =   3675
      Begin VB.OptionButton optSize 
         Caption         =   "&24 x 24"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   195
         TabIndex        =   17
         Top             =   645
         Width           =   1245
      End
      Begin VB.OptionButton optSize 
         Caption         =   "&Custom:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1290
         TabIndex        =   8
         Top             =   345
         Width           =   930
      End
      Begin VB.OptionButton optSize 
         Caption         =   "&48 x 48"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   195
         TabIndex        =   7
         Top             =   1260
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton optSize 
         Caption         =   "&32 x 32"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   195
         TabIndex        =   6
         Top             =   955
         Width           =   1245
      End
      Begin VB.OptionButton optSize 
         Caption         =   "&16 x 16"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Top             =   345
         Width           =   1245
      End
      Begin iICO.ucUpDownBox ucWidth 
         Height          =   315
         Left            =   2790
         TabIndex        =   10
         Top             =   285
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Enabled         =   0   'False
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
      Begin iICO.ucUpDownBox ucHeight 
         Height          =   315
         Left            =   2790
         TabIndex        =   12
         Top             =   645
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Enabled         =   0   'False
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
      Begin VB.Label lblHeight 
         Caption         =   "Height"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2265
         TabIndex        =   11
         Top             =   705
         Width           =   615
      End
      Begin VB.Label lblWidth 
         Caption         =   "Width"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2265
         TabIndex        =   9
         Top             =   345
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3375
      Width           =   1050
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3375
      Width           =   1050
   End
   Begin VB.Frame fraColorDepth 
      Caption         =   "Color depth"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   135
      TabIndex        =   13
      Top             =   150
      Width           =   3675
      Begin VB.OptionButton optBPP 
         Caption         =   "&ARGB color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1590
         TabIndex        =   4
         Top             =   630
         Width           =   1455
      End
      Begin VB.OptionButton optBPP 
         Caption         =   "&True color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1590
         TabIndex        =   3
         Top             =   315
         Width           =   1455
      End
      Begin VB.OptionButton optBPP 
         Caption         =   "&256 colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   195
         TabIndex        =   2
         Top             =   945
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optBPP 
         Caption         =   "&16 colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   195
         TabIndex        =   1
         Top             =   630
         Width           =   1455
      End
      Begin VB.OptionButton optBPP 
         Caption         =   "&Monochrome"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   195
         TabIndex        =   0
         Top             =   315
         Width           =   1455
      End
   End
End
Attribute VB_Name = "fNewFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAXSIZE As Long = 64
Private m_Cancel      As Boolean

'//

Public Sub SetDialogMode(ByVal ResourceType As ResourceTypeCts)

  Dim nOpt As Integer

    Select Case ResourceType
        Case [rtIcon]   '-- Icon
            For nOpt = 0 To 4
                optSize(nOpt).Enabled = True
            Next nOpt
        Case [rtCursor] '-- Cursor
            For nOpt = 0 To 4
                If (nOpt <> 2) Then
                    optSize(nOpt).Enabled = False
                End If
            optSize(2) = True
        Next nOpt
    End Select
End Sub

Private Sub Form_Load()
    
    Set Me.Icon = Nothing
    
    Call mMisc.RemoveButtonBorderEnhance(Me.cmdOk)
    Call mMisc.RemoveButtonBorderEnhance(Me.cmdCancel)
    
    With ucWidth
        .Min = 1
        .Max = MAXSIZE
        .Value = 16
    End With
    With ucHeight
        .Min = 1
        .Max = MAXSIZE
        .Value = 16
    End With
End Sub

'//

Private Sub optSize_Click(Index As Integer)
    
  Dim bEnable As Boolean
  
    bEnable = optSize(4).Value
    
    lblWidth.Enabled = bEnable
    ucWidth.Enabled = bEnable
    
    lblHeight.Enabled = bEnable
    ucHeight.Enabled = bEnable
    
    If (Index = 4) Then
        ucWidth.SetFocus
    End If
End Sub

'//

Private Sub cmdOk_Click()
    m_Cancel = False
    Call Me.Hide
End Sub

Private Sub cmdCancel_Click()
    m_Cancel = True
    Call Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '-- Close button pressed
    If (UnloadMode = vbFormControlMenu) Then
        Call cmdCancel_Click
    End If
End Sub

Public Property Get Cancel() As Boolean
    Cancel = m_Cancel
End Property


Public Property Get NewWidth() As Long
    
    Select Case True
        Case optSize(0): NewWidth = 16
        Case optSize(1): NewWidth = 24
        Case optSize(2): NewWidth = 32
        Case optSize(3): NewWidth = 48
        Case optSize(4): NewWidth = ucWidth.Value
    End Select
End Property

Public Property Get NewHeight() As Long
    
    Select Case True
        Case optSize(0): NewHeight = 16
        Case optSize(1): NewHeight = 24
        Case optSize(2): NewHeight = 32
        Case optSize(3): NewHeight = 48
        Case optSize(4): NewHeight = ucHeight.Value
    End Select
End Property

Public Property Get NewBPP() As Long
    
    Select Case True
        Case optBPP(0): NewBPP = 1
        Case optBPP(1): NewBPP = 4
        Case optBPP(2): NewBPP = 8
        Case optBPP(3): NewBPP = 24
        Case optBPP(4): NewBPP = 32
    End Select
End Property
