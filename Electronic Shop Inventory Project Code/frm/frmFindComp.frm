VERSION 5.00
Begin VB.Form frmFindComp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find component"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindComp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkFilter 
      Appearance      =   0  'Flat
      Caption         =   "Whole Word"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3300
      TabIndex        =   7
      Tag             =   "1"
      Top             =   960
      Width           =   1275
   End
   Begin VB.CheckBox chkFilter 
      Appearance      =   0  'Flat
      Caption         =   "Match Case"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3300
      TabIndex        =   8
      Tag             =   "2"
      Top             =   1260
      Width           =   1935
   End
   Begin VB.Frame Frame 
      Caption         =   "Search for"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   3075
      Begin VB.OptionButton obtnField 
         Appearance      =   0  'Flat
         Caption         =   "Box Number"
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
         Left            =   1560
         TabIndex        =   6
         Tag             =   "7"
         Top             =   600
         Width           =   1215
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
         Left            =   1560
         TabIndex        =   5
         Tag             =   "2"
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton obtnField 
         Appearance      =   0  'Flat
         Caption         =   "Code"
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
         Tag             =   "1"
         Top             =   600
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton obtnField 
         Appearance      =   0  'Flat
         Caption         =   "Product Name"
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
         Tag             =   "0"
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4260
      TabIndex        =   10
      Top             =   540
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Height          =   315
      Left            =   4260
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Caption         =   "Find What:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "frmFindComp"
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

Private Sub cmdFind_Click()

  If OldFind = NewFind And NewFind > 0 Then         '== Ovo nam je potrebno ako pretrazimo
    NewFind = 0                                     '== cijelu tabelu da se vratimo na pocetak
    OldFind = 0                                     '== i da kruzimo
    Call Find_String(Trim(txtSearch.Text), NewFind)
    Exit Sub
  End If

  OldFind = NewFind
  Call Find_String(Trim(txtSearch.Text), NewFind)

End Sub

Private Sub Find_String(sText As String, iStart As Integer)
  
  Dim lCol              As Long       '== Broj kolone
  Dim lRow              As Long       '== Broj reda
  Dim bFound            As Boolean    '== Da znamo kad smo nasli podatak
  Dim i                 As Integer    '== Brojac
  Dim iFilter           As Integer    '== Za MatchCase i WholeWord filter

  iFilter = Set_Filter
  
  For i = 0 To 3
    If obtnField(i).value = True Then
      lCol = CLng(obtnField(i).Tag)
    End If
  Next i

  For lRow = iStart To frmComp.FlexGrid.Rows - 1
  
    If iFilter = 0 Then
    
      If InStr(Trim(LCase(frmComp.FlexGrid.TextMatrix(lRow, lCol))), Trim(LCase(sText))) Then
        bFound = True
        NewFind = lRow + 1
        Sel_Row (lRow)  '== bojanje reda
        lRow = frmComp.FlexGrid.Rows - 1  '== Izlaz iz rutine (umjesto Exit Sub)
      End If
      
    ElseIf iFilter = 1 Then

      If Trim(LCase(frmComp.FlexGrid.TextMatrix(lRow, lCol))) = Trim(LCase(sText)) Then
        bFound = True
        NewFind = lRow + 1
        Sel_Row (lRow)
        lRow = frmComp.FlexGrid.Rows - 1
      End If
      
    ElseIf iFilter = 2 Then
    
      If InStr(Trim(frmComp.FlexGrid.TextMatrix(lRow, lCol)), Trim(sText)) Then
        bFound = True
        NewFind = lRow + 1
        Sel_Row (lRow)
        lRow = frmComp.FlexGrid.Rows - 1
      End If
      
    ElseIf iFilter = 3 Then
    
      If Trim(frmComp.FlexGrid.TextMatrix(lRow, lCol)) = Trim(sText) Then
        bFound = True
        NewFind = lRow + 1
        Sel_Row (lRow)
        lRow = frmComp.FlexGrid.Rows - 1
      End If

    End If
      
  Next lRow
  
End Sub

Private Function Set_Filter() As Integer
  
  Dim i As Integer
  
  Set_Filter = 0
  
  For i = 0 To 1
    If chkFilter(i).value Then
      Set_Filter = Set_Filter + CInt(chkFilter(i).Tag)
    End If
  Next i
  
End Function

Private Sub Sel_Row(Selected As Long)

    frmComp.FlexGrid.Row = Selected
    frmComp.FlexGrid.RowSel = Selected
    frmComp.FlexGrid.Col = 0
    frmComp.FlexGrid.ColSel = frmComp.FlexGrid.Cols - 1
    
    If Selected > 20 Then
      frmComp.FlexGrid.TopRow = Selected
    End If
    
End Sub

