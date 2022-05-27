VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4485
   ClientLeft      =   2340
   ClientTop       =   1650
   ClientWidth     =   4320
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Height          =   4515
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   1335
      Begin VB.Image Image 
         Height          =   4350
         Left            =   60
         Picture         =   "frmAbout.frx":000C
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3060
      TabIndex        =   0
      Top             =   4020
      Width           =   1155
   End
   Begin VB.Label lblVer 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   4080
      Width           =   915
   End
   Begin VB.Line Line 
      Index           =   11
      X1              =   1380
      X2              =   4260
      Y1              =   3675
      Y2              =   3675
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   10
      X1              =   1380
      X2              =   4260
      Y1              =   3660
      Y2              =   3660
   End
   Begin VB.Line Line 
      Index           =   9
      X1              =   1380
      X2              =   4260
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   8
      X1              =   1380
      X2              =   4260
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line 
      Index           =   7
      X1              =   1380
      X2              =   4260
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   1380
      X2              =   4260
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line 
      Index           =   5
      X1              =   1380
      X2              =   4260
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   1380
      X2              =   4260
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   1380
      X2              =   4260
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line Line 
      Index           =   2
      X1              =   1380
      X2              =   4260
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmAbout.frx":344E
      Height          =   1035
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Top             =   1260
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Caption         =   "e-mail: rz@edaboard.com"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Top             =   3180
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Caption         =   "Author: Refik Zaimoviæ"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1380
      X2              =   4260
      Y1              =   1060
      Y2              =   1060
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   1380
      X2              =   4260
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblInfo 
      Caption         =   "Service and Workshop Inventory Application"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Index           =   0
      Left            =   1500
      TabIndex        =   2
      Top             =   240
      Width           =   2715
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()

  Unload Me
  
End Sub

Private Sub Form_Load()
  
  SetWindowLong cmdExit.hwnd, GWL_EXSTYLE, 131076
  SetWindowPos cmdExit.hwnd, 0, 0, 0, 0, 0, wFLAGS
  
  lblVer.Caption = "(Ver. " & App.Major & "." & App.Minor & "." & App.Revision & ")"
  
  MLSLoadLanguage Me

End Sub


' -----------------------------------------------------------------
' This module was made by Multi-Language Support Add-in for VB,
' by Giorgio Brausi (gibra)
' Contact me by e-mail: vbcorner@lycos.it or gibra@amc2000.it
' Web site: http://utenti.lycos.it/vbcorner/
