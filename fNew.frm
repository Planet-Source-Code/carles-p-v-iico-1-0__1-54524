VERSION 5.00
Begin VB.Form fNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New resource"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
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
   ScaleHeight     =   88
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   206
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraType 
      Caption         =   "Resource type"
      Height          =   960
      Left            =   135
      TabIndex        =   2
      Top             =   150
      Width           =   1530
      Begin VB.OptionButton optType 
         Caption         =   "Cursor"
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   4
         Top             =   585
         Width           =   1170
      End
      Begin VB.OptionButton optType 
         Caption         =   "Icon"
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   285
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1050
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1050
   End
End
Attribute VB_Name = "fNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Cancel As Boolean

'//

Private Sub Form_Load()
    
    Set Me.Icon = Nothing

    Call mMisc.RemoveButtonBorderEnhance(Me.cmdOk)
    Call mMisc.RemoveButtonBorderEnhance(Me.cmdCancel)
End Sub

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

'//

Public Property Get Cancel() As Boolean
    Cancel = m_Cancel
End Property

Public Property Get ResourceType() As ResourceTypeCts
    ResourceType = IIf(optType(0), 1, 2)
End Property

