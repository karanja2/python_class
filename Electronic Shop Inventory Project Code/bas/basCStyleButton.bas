Attribute VB_Name = "basCStyleButton"
Option Explicit

Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" _
                          (ByVal hwnd As Long, _
                          ByVal nIndex As Long, _
                          ByVal dwNewLong As Long)

Public Declare Function SetWindowPos& Lib "user32" _
                          (ByVal hwnd As Long, _
                          ByVal hWndInsertAfter As Long, _
                          ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal cx As Long, _
                          ByVal cy As Long, _
                          ByVal wFLAGS As Long)

Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200

Public Const wFLAGS = _
    SWP_NOMOVE Or _
    SWP_NOSIZE Or _
    SWP_NOOWNERZORDER Or _
    SWP_NOZORDER Or _
    SWP_FRAMECHANGED

Public Const GWL_EXSTYLE = (-20)
