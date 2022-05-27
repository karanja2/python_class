Attribute VB_Name = "UseIni"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Function GetProfile(lpAppName$, lpKeyName$, lpDefault, lpFileName$)

    Dim lpReturnString$, nSize%, Valid%
    
    lpReturnString$ = Space$(128)
    nSize% = Len(lpReturnString$)
    Valid% = GetPrivateProfileString(ByVal lpAppName$, _
                                     ByVal lpKeyName$, _
                                     ByVal lpDefault, _
                                     ByVal lpReturnString$, _
                                     ByVal nSize%, _
                                     ByVal lpFileName$)
    GetProfile = Left$(lpReturnString$, Valid%)

End Function

Sub WriteProfile(lpAppName$, lpKeyName$, lpString$, lpFileName$)
    
    Dim Valid%
    
    Valid% = WritePrivateProfileString(lpAppName$, lpKeyName$, lpString$, lpFileName$)

End Sub

Function GetProfileSection(lpAppName As String, lpFileName As String) As String
    
    Dim strReturnString As String
    Dim lSize As Long, lValid As Long
    
    strReturnString = Space$(256)
    lSize = Len(strReturnString)
    lValid = GetPrivateProfileSection(ByVal lpAppName, _
                                      ByVal strReturnString, _
                                      ByVal lSize, ByVal lpFileName)
        GetProfileSection = Left$(strReturnString, lValid)

End Function
