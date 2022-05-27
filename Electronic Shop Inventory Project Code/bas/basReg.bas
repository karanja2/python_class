Attribute VB_Name = "basReg"
''--------------------------------------------------------------------------------
''Za pisanje u registar koristi se:
''SetRegValue HKEY_LOCAL_MACHINE,
'''"Software\Microsoft\Windows\CurrentVersion\Run", "proba", App.Path + "\ime.exe"
''
''Za brisanje iz registra koristi se:
''DeleteValue HKEY_LOCAL_MACHINE,
'''"Software\Microsoft\Windows\CurrentVersion\Run", "proba"
''--------------------------------------------------------------------------------
Public Type SECURITY_ATTRIBUTES
  nLength                                  As Long
  lpSecurityDescriptor                     As Long
  bInheritHandle                           As Long
End Type
Public Enum T_KeyClasses
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
End Enum
#If False Then
Private HKEY_CLASSES_ROOT, HKEY_CURRENT_CONFIG, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS
#End If
Private Const SYNCHRONIZE                  As Long = &H100000
Private Const STANDARD_RIGHTS_ALL          As Long = &H1F0000
Private Const KEY_QUERY_VALUE              As Long = &H1
Private Const KEY_SET_VALUE                As Long = &H2
Private Const KEY_CREATE_LINK              As Long = &H20
Private Const KEY_CREATE_SUB_KEY           As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS       As Long = &H8
Private Const KEY_NOTIFY                   As Long = &H10
Private Const KEY_ALL_ACCESS               As Double = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const REG_DWORD                    As Integer = 4
Private Const REG_SZ                       As Integer = 1
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal lpReserved As Long, _
                                                                                  ByRef lpType As Long, _
                                                                                  ByVal lpData As String, _
                                                                                  ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal Reserved As Long, _
                                                                                  ByVal dwType As Long, _
                                                                                  ByVal lpData As String, _
                                                                                  ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                                                                    ByVal lpValueName As String) As Long

Public Sub DeleteValue(rClass As T_KeyClasses, _
                       Path As String, _
                       sKey As String)

  Dim hKey As Long
  Dim res  As Long

  res = RegOpenKeyEx(rClass, Path, 0, KEY_ALL_ACCESS, hKey)
  res = RegDeleteValue(hKey, sKey)
  RegCloseKey hKey

End Sub

Public Function SetRegValue(KeyRoot As T_KeyClasses, _
                            Path As String, _
                            sKey As String, _
                            NewValue As String) As Boolean

  Dim hKey       As Long
  Dim KeyValType As Long
  Dim KeyValSize As Long
  Dim tmpVal     As String
  Dim res        As Long
  Dim i          As Integer
  Dim X          As Long

  res = RegOpenKeyEx(KeyRoot, Path, 0, KEY_ALL_ACCESS, hKey)
  If res <> 0 Then
    GoTo Errore
  End If
  tmpVal = String$(1024, 0)
  KeyValSize = 1024
  res = RegQueryValueEx(hKey, sKey, 0, KeyValType, tmpVal, KeyValSize)
  Select Case res
   Case 2
    KeyValType = REG_SZ
   Case Is <> 0
    GoTo Errore
  End Select
  Select Case KeyValType
   Case REG_SZ
    tmpVal = NewValue
   Case REG_DWORD
    X = Val(NewValue)
    tmpVal = vbNullString
    For i = 0 To 3
      tmpVal = tmpVal & Chr$(X Mod 256)
      X = X \ 256
    Next i
  End Select
  KeyValSize = Len(tmpVal)
  res = RegSetValueEx(hKey, sKey, 0, KeyValType, tmpVal, KeyValSize)
  If res <> 0 Then
    GoTo Errore
  End If
  SetRegValue = True
  RegCloseKey hKey

Exit Function

Errore:
  SetRegValue = False
  RegCloseKey hKey

End Function


