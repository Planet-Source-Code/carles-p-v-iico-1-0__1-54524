Attribute VB_Name = "mSettings"
Option Explicit
Option Compare Text

'-- API :

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'//

Public Sub LoadSettings()

    '-- Main form
    With fMain
        .Width = pvGetINI("iICO.ini", "Forms", "MainWidth", .Width)
        .Height = pvGetINI("iICO.ini", "Forms", "MainHeight", .Height)
        .Top = pvGetINI("iICO.ini", "Forms", "MainTop", (Screen.Height - .Height) \ 2)
        .Left = pvGetINI("iICO.ini", "Forms", "MainLeft", (Screen.Width - .Width) \ 2)
        .WindowState = pvGetINI("iICO.ini", "Forms", "MainWindowState", .WindowState)
    End With
    
    '-- Color
    G_ColorScreen = pvGetINI("iICO.ini", "Color", "ColorScreen", &HFAC896)

    '-- Paths
    G_OpenSavePath = pvGetINI("iICO.ini", "Paths", "OpenSavePath", vbNullString)
    G_ImportPath = pvGetINI("iICO.ini", "Paths", "ImportPath", vbNullString)
    G_PalettePath = pvGetINI("iICO.ini", "Paths", "PalettePath", vbNullString)
End Sub

Public Sub SaveSettings()
    
    '-- Main form
    With fMain
        If (.WindowState = vbNormal) Then
            Call pvPutINI("iICO.ini", "Forms", "MainWidth", .Width)
            Call pvPutINI("iICO.ini", "Forms", "MainHeight", .Height)
            Call pvPutINI("iICO.ini", "Forms", "MainTop", .Top)
            Call pvPutINI("iICO.ini", "Forms", "MainLeft", .Left)
        End If
        Call pvPutINI("iICO.ini", "Forms", "MainWindowState", .WindowState)
    End With
    
    '-- Color
    Call pvPutINI("iICO.ini", "Color", "ColorScreen", CStr(G_ColorScreen))

    '-- Paths
    Call pvPutINI("iICO.ini", "Paths", "OpenSavePath", G_OpenSavePath)
    Call pvPutINI("iICO.ini", "Paths", "ImportPath", G_ImportPath)
    Call pvPutINI("iICO.ini", "Paths", "PalettePath", G_PalettePath)
End Sub
    
'//

Private Function pvAppPath() As String

    pvAppPath = App.Path & IIf(Right$(App.Path, 1) = "\", vbNullString, "\")
End Function

Private Sub pvPutINI(INIFile As String, INIHead As String, INIKey As String, INIVal As String)
  
  Dim INIFileName As String
  Dim sRet        As String
  
    INIFileName = pvAppPath & INIFile
    sRet = WritePrivateProfileString(INIHead, INIKey, INIVal, INIFileName)
End Sub

Private Function pvGetINI(INIFile As String, INIHead As String, INIKey As String, INIDefault As String) As String

  Dim INIFileName As String
  Dim Temp        As String * 260
  Dim sRet        As String
    
    INIFileName = pvAppPath & INIFile
    sRet = GetPrivateProfileString(INIHead, INIKey, INIDefault, Temp, Len(Temp), INIFileName)
    pvGetINI = Trim$(Temp)
    pvGetINI = Left$(pvGetINI, Len(pvGetINI) - 1)
End Function


