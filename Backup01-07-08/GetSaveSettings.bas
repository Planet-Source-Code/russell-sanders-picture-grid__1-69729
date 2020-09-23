Attribute VB_Name = "GetSaveSettings"
' This is a replacement for vbs' GetSetting SaveSetting functions useing initial files Must pass peram as string

Option Explicit

Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetSetting(AppName As String, Section As String, Setting As String, Default As Variant) As Variant
Dim Ret As String 'set a veriable to contain a buffer
Dim Pos As Long   'will hold the length of the returned string
    Ret = String(255, 0)    'Create the buffer
    Pos = GetPrivateProfileString(Section, Setting, Default, Ret, 255, App.Path & "\" & AppName & ".int") 'retrive the setting
    GetSetting = Left(Ret, Pos) 'trim the returned string to use
        If Len(GetSetting) <= 0 Then
            GetSetting = Default
        End If
End Function

Public Sub SaveSetting(AppName, Section As String, Key As String, Setting As String)
    WritePrivateProfileString Section, Key, Setting, App.Path & "\" & AppName & ".int"  'save the setting
End Sub

