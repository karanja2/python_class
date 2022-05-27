VERSION 5.00
Begin VB.Form frmFindRecord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find record"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindRecord.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Next"
      Height          =   315
      Left            =   4320
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4320
      TabIndex        =   10
      Top             =   540
      Width           =   1095
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   2715
   End
   Begin VB.Frame Frame 
      Caption         =   "Search for"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   3015
      Begin VB.OptionButton obtnField 
         Appearance      =   0  'Flat
         Caption         =   "Worker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Tag             =   "Worker"
         Top             =   300
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton obtnField 
         Appearance      =   0  'Flat
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Tag             =   "Date_"
         Top             =   600
         Width           =   1035
      End
      Begin VB.OptionButton obtnField 
         Appearance      =   0  'Flat
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1260
         TabIndex        =   5
         Tag             =   "Description"
         Top             =   300
         Width           =   1335
      End
      Begin VB.OptionButton obtnField 
         Appearance      =   0  'Flat
         Caption         =   "Customer Info"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1260
         TabIndex        =   6
         Tag             =   "Customer_Info"
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.CheckBox chkFilter 
      Appearance      =   0  'Flat
      Caption         =   "Match Case"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      Tag             =   "2"
      Top             =   1260
      Width           =   2115
   End
   Begin VB.CheckBox chkFilter 
      Appearance      =   0  'Flat
      Caption         =   "Whole Word"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   7
      Tag             =   "1"
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label lblInfo 
      Caption         =   "Find What:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1155
   End
End
Attribute VB_Name = "frmFindRecord"
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

Private OldFind                   As Integer    '== Cuva prethodni rezultat pretrage
Private NewFind                   As Integer    '== Za pretragu

Private Sub cmdExit_Click()

  Unload Me
  
End Sub

Private Sub cmdFind_Click()

  If OldFind = NewFind And NewFind > 0 Then         '== Ovo nam je potrebno ako pretrazimo
    NewFind = 0                                     '== cijelu bazu da se vratimo na pocetak
    OldFind = 0                                     '== i da kruzimo po bazi
    Call Find_String(Trim(txtSearch.Text), NewFind)
    Exit Sub
  End If

  OldFind = NewFind
  Call Find_String(Trim(txtSearch.Text), NewFind)

End Sub

Private Sub Find_String(TextSearch As String, _
                        Start As Integer)                '== pretrazujemo RS

  Dim Found             As Boolean    '== Da znamo kad smo nasli podatak
  Dim i                 As Integer    '== Brojac
  Dim sField            As String     '== Za Fields (Nazivi polja se nalaze u tagu option dugmeta
  Dim iFilter           As Integer    '== Za MatchCase i WholeWord filter

  iFilter = Set_Filter
  
  For i = 0 To 3
    If obtnField(i).value Then
      sField = obtnField(i).Tag
    End If
  Next i

  RS2.Move NewFind, 1
  With RS2
    For i = Start To .RecordCount - 1
        If .Fields(sField) <> "" Then
          Select Case iFilter
            Case 0
              If InStr(Trim(LCase(.Fields(sField))), Trim(LCase(TextSearch))) > 0 Then
                Found = True
                NewFind = i + 1
                .Move i, 1
                frmRecord.Populate_Form
                i = .RecordCount - 1  '== prije je na ovom mjesu islo Exit Sub ali posto explicini izlaz iz rutine nije dobro rjesenje postavili smo brojac na kraj i tako izasli regularno iz ove rutine.
              End If                  '== i dalje ostaje da se ova pretraga malo bolje rijesi
            Case 1
              If Trim(LCase(.Fields(sField))) = Trim(LCase(TextSearch)) Then
                Found = True
                NewFind = i + 1
                .Move i, 1
                frmRecord.Populate_Form
                i = .RecordCount - 1
              End If
            Case 2
              If InStr(Trim(.Fields(sField)), Trim(TextSearch)) > 0 Then
                Found = True
                NewFind = i + 1
                .Move i, 1
                frmRecord.Populate_Form
                i = .RecordCount - 1
              End If
            Case 3
              If Trim(.Fields(sField)) = Trim(TextSearch) Then
                Found = True
                NewFind = i + 1
                RS2.Move i, 1
                frmRecord.Populate_Form
                i = .RecordCount - 1
              End If
          End Select
        End If
      .MoveNext
    Next i
  End With
  
End Sub


Private Sub Form_Load()

  SetWindowLong cmdFind.hwnd, GWL_EXSTYLE, 131076
  SetWindowPos cmdFind.hwnd, 0, 0, 0, 0, 0, wFLAGS
  SetWindowLong cmdExit.hwnd, GWL_EXSTYLE, 131076
  SetWindowPos cmdExit.hwnd, 0, 0, 0, 0, 0, wFLAGS
  
  MLSLoadLanguage Me

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

Private Function Set_Filter() As Integer
'== Provjeravamo koji filteri (Match Case i Whole Word) su ukljuceni. Koristeno je tag svojstvo
  Dim i As Integer
  
  Set_Filter = 0
  
  For i = 0 To 1
    If chkFilter(i).value Then
      Set_Filter = Set_Filter + CInt(chkFilter(i).Tag)
    End If
  Next i
  
End Function
