Attribute VB_Name = "basSysTray"
Option Explicit

  Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

  Public Const NIM_ADD                As Long = &H0
  Private Const NIM_MODIFY            As Long = &H1
  Public Const NIM_DELETE             As Long = &H2
  Public Const NIF_MESSAGE            As Long = &H1
  Public Const NIF_ICON               As Long = &H2
  Public Const NIF_TIP                As Long = &H4
  Public Const WM_MOUSEMOVE           As Long = &H200
  
  Private Const WM_LBUTTONDOWN        As Long = &H201
  Private Const WM_LBUTTONUP          As Long = &H202
  
  Public Const WM_LBUTTONDBLCLK       As Long = &H203
  Public Const WM_RBUTTONDOWN         As Long = &H204
  
  Private Const WM_RBUTTONUP          As Long = &H205
  Private Const WM_RBUTTONDBLCLK      As Long = &H206
  Private Const HWND_TOPMOST          As Integer = -1
  
  Public nid                          As NOTIFYICONDATA
  
  Public Type NOTIFYICONDATA
    cbSize                            As Long
    hwnd                              As Long
    uId                               As Long
    uFlags                            As Long
    uCallBackMessage                  As Long
    hIcon                             As Long
    szTip                             As String * 64
  End Type
  


