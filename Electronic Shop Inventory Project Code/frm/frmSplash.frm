VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3885
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   2000
      Left            =   7080
      Top             =   300
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail: rz@edaboard.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   3600
      Width           =   2475
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(c) Refik Zaimovic - 2009"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   1
      Top             =   3360
      Width           =   2475
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Electronic Workshop Inventory Application"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   6975
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                               ByVal hWndInsertAfter As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               ByVal cx As Long, _
                                               ByVal cy As Long, _
                                               ByVal wFLAGS As Long)

Private Const HWND_TOPMOST        As Integer = -1
Private Const HWND_NOTOPMOST      As Integer = -2
Private Const SWP_NOSIZE          As Long = &H1
Private Const SWP_NOMOVE          As Long = &H2
Private Const SWP_NOACTIVATE      As Long = &H10
Private Const SWP_SHOWWINDOW      As Long = &H40
Private TopMost                   As Boolean

Private Sub Form_Load()

  TopMost = True
  SetTopMost

End Sub

Private Sub Form_Unload(Cancel As Integer)

  TopMost = False
  SetTopMost
  
End Sub

Private Sub SetTopMost()

  If TopMost Then
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
   Else
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
  End If

End Sub

Private Sub Timer_Timer()

  On Error Resume Next
  Timer.Enabled = False
  Unload Me
  Set frmSplash = Nothing

End Sub
''******************************************************************************************************
''Jedan nacin koriscenja splash screen-a je sa tajmerom. Splash se ucitava na pocetku, ali nije modalan
''kao sada, i dopusta ucitavanje ostalih formi. Redosled ucitavanja formi je sledeci: Prvo se ucita
''splash forma (non modal), pa onda (iza splasha) glavna forma sa svim ostalim
''child formama ako su setovane da se ucitaju, i na kraju Login forma koja je
''modalna. U basMain modulu u Main rutini treba postaviti da splash forma NIJE modalna.
''Tada se koristi deklaracija na vrhu koja setuje formu splash kao TopMost i
''treba omoguciti ucitavanje TopMost na form_load i ponistavanje na Form_Unload.
''******************************************************************************************************
''Drugi nacin je da podesimo splash formu na VbModal u basMain modulu u Main rutini.
''Tada stopiramo ucitavanje programa dok ne ugasimo splash formu (sa klikom ili ESC),
''Ovaj nacin ima taj nedostatak ako imamo veliku bazu koja se duze ucitava, posle
''gasenja splash forme javice se pauza dok se podaci ne ucitaju (ako smo podesili
''child forme pale sa glavnom formom). U ovom slucaju ne koristimo TopMost deklaraciju\
''na vrhu, ne setujemo TomMost na Form_Load i Form_Unload i nemamo tajmera.
''******************************************************************************************************

