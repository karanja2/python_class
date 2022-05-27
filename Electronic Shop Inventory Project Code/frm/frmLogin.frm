VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login to application"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4095
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2460
      TabIndex        =   4
      Top             =   1140
      Width           =   1275
   End
   Begin VB.CommandButton cmdDelLog 
      Caption         =   "Delete .log"
      Height          =   375
      Left            =   300
      TabIndex        =   3
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   60
      X2              =   4020
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   60
      X2              =   4020
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label lblUser 
      Caption         =   "User Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "User name must be minimum 3 characters long"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oTextStream As Scripting.TextStream

Private Sub cmdDelLog_Click()

  If LenB(Dir(App.Path & "\userlog.tfl")) <> 0 Then
    Kill App.Path & "\userlog.tfl"
  End If
  
End Sub

Private Sub cmdOk_Click()

  sWorker = txtUser.Text
  
  If Len(sWorker) > 2 Then
  
    Set oFSO = New Scripting.FileSystemObject
        
    Set oTextStream = oFSO.OpenTextFile(App.Path & "\userlog.tfl", ForAppending, True, TristateFalse)
    oTextStream.WriteLine (CStr(Date) & " User loged = '" & sWorker & "'")
    
    oTextStream.Close
    Set oTextStream = Nothing
    Set oFSO = Nothing

    WriteProfile "LastUser", "Value", txtUser.Text, App.Path & "\service.ini"
    MDI_Main.SBar.Panels(2).Text = MLSGetString("0032") & sWorker & ")"
    Unload Me
    
  Else
  
    lblInfo.Visible = True
    
  End If
  
End Sub

Private Sub Form_Load()

  SetWindowLong cmdDelLog.hwnd, GWL_EXSTYLE, 131076
  SetWindowPos cmdDelLog.hwnd, 0, 0, 0, 0, 0, wFLAGS
  SetWindowLong cmdOk.hwnd, GWL_EXSTYLE, 131076
  SetWindowPos cmdOk.hwnd, 0, 0, 0, 0, 0, wFLAGS
  
  MLSLoadLanguage Me
  txtUser.Text = GetProfile("LastUser", "Value", "0", App.Path & "\service.ini")
  txtUser.SelStart = 0
  txtUser.SelLength = (Len(txtUser.Text))
  
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)

  If KeyAscii = 44 Then
    KeyAscii = 0
  End If
  
End Sub
