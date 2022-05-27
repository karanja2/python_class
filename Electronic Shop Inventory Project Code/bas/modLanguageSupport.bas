Attribute VB_Name = "basLanguage"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub MLSLoadLanguage(Form As Form)

    Dim obj As Object
    Dim sFileName As String
    Dim a As String
    
    On Error Resume Next
    
    If Right(App.Path, 1) = "\" Then
    
        sFileName = App.Path & sLaguage & ".lng"
        
    Else
    
        sFileName = App.Path & "\" & sLaguage & ".lng"
        
    End If

    If Len(Form.Caption) > 0 Then
    
        Form.Caption = MLSReadINI(sFileName, CStr(Form.Name), CStr(Form.Name) & ".Caption")
    
    End If

    Form.ToolTipText = MLSReadINI(sFileName, CStr(Form.Name), CStr(Form.Name) & ".ToolTipText")
    Form.Tag = MLSReadINI(sFileName, CStr(Form.Name), CStr(Form.Name) & ".Tag")

    For Each obj In Form
    
        Dim bHasIndex As Boolean
        
        a$ = ""
        
        bHasIndex = (obj.Index >= 0)
        
        If Err.Number = 343 Then
        
            bHasIndex = False
            Err.Clear
            
        End If

        If bHasIndex Then
        
            a$ = MLSReadINI(sFileName, CStr(Form.Name), obj.Name & "(" & obj.Index & ").Caption")
        
        Else
            
            a$ = MLSReadINI(sFileName, CStr(Form.Name), obj.Name & ".Caption")
        
        End If
        
        If a$ <> "" Then
            
            obj.Caption = a$
        
        End If

        If obj.Tabs Then
        
            If Err = 0 Then
                a$ = ""
                Dim nT As Integer
                For nT = 0 To obj.Tabs
                     obj.TabCaption(nT) = MLSReadINI(sFileName, CStr(Form.Name), obj.Name & ".TabCaption(" & nT & ")")
                Next nT
            Else
            
                Err.Clear
                
            End If
            
        End If

        DoEvents

    Next
End Sub

Public Function MLSGetString(KeyName As String) As String

Dim sFileName As String

    On Error Resume Next
    
    If sLaguage = "" Then
    
        sLaguage = MLSReadINI(App.Path & "\" & "LangSetting.ini", "Language", "CurrentLanguage")
    
    End If
    
    If Right(App.Path, 1) = "\" Then
        
        sFileName = App.Path & sLaguage & ".lng"
    
    Else
        
        sFileName = App.Path & "\" & sLaguage & ".lng"
    
    End If
    
    MLSGetString = MLSReadINI(sFileName, "Strings", KeyName$)

End Function
Public Function MLSReadINI(file$, SectionName$, KeyName$) As String

Dim value As String * 1024
Dim i As Long

  i = GetPrivateProfileString(SectionName$, KeyName$, "", value, 512, file$)
  MLSReadINI = Left$(value, InStr(value, Chr$(0)) - 1)
        
End Function
Public Function MLSWriteINI(file$, SectionName$, KeyName$, NewValue$) As Long
    
    MLSWriteINI = WritePrivateProfileString(SectionName$, KeyName$, NewValue$, file$)

End Function
