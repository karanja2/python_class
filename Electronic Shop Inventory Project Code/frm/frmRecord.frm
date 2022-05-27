VERSION 5.00
Begin VB.Form frmRecord 
   Caption         =   "Records"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecord.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   5895
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4500
      TabIndex        =   18
      Top             =   5940
      Width           =   1275
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3000
      TabIndex        =   17
      Top             =   5940
      Width           =   495
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2460
      TabIndex        =   16
      Top             =   5940
      Width           =   495
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   15
      Top             =   5940
      Width           =   495
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1380
      TabIndex        =   14
      Top             =   5940
      Width           =   495
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   3840
      TabIndex        =   21
      Tag             =   "No"
      Top             =   5940
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblRecCount 
      Height          =   255
      Left            =   1260
      TabIndex        =   19
      Top             =   5160
      Width           =   4515
   End
   Begin VB.Line ln2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   5880
      Y1              =   5595
      Y2              =   5595
   End
   Begin VB.Line ln1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   5880
      Y1              =   5580
      Y2              =   5580
   End
   Begin VB.Label lblTrue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Yes"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5220
      TabIndex        =   5
      Tag             =   "Charged"
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label lblData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1515
      Index           =   5
      Left            =   1260
      TabIndex        =   13
      Tag             =   "Customer_Info"
      Top             =   3540
      Width           =   4515
   End
   Begin VB.Label lblData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1515
      Index           =   4
      Left            =   1260
      TabIndex        =   12
      Tag             =   "Description"
      Top             =   1920
      Width           =   4515
   End
   Begin VB.Label lblData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   1260
      TabIndex        =   11
      Tag             =   "Parts_Sold"
      Top             =   1500
      Width           =   4515
   End
   Begin VB.Label lblData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   1260
      TabIndex        =   4
      Tag             =   "Price"
      Top             =   1080
      Width           =   2235
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Charged:"
      Height          =   255
      Index           =   6
      Left            =   4020
      TabIndex        =   10
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer Info:"
      Height          =   495
      Index           =   5
      Left            =   60
      TabIndex        =   9
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Description:"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   8
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Parts Sold:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Price:"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label lblData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   1260
      TabIndex        =   2
      Tag             =   "Date_"
      Top             =   660
      Width           =   2235
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   1260
      TabIndex        =   1
      Tag             =   "Worker"
      Top             =   240
      Width           =   2235
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Worker:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   1095
   End
   Begin VB.Label lblFalse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5220
      TabIndex        =   20
      Top             =   1080
      Width           =   555
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

  Conn_Close
  Unload Me
  
End Sub

Private Sub cmdNavi_Click(Index As Integer)

  Select Case Index
  
    Case 0:                                 '== move first
      If Not RS2.BOF Then RS2.MoveFirst
    
    Case 1                                  '== move previous
      If Not RS2.BOF Then RS2.MovePrevious
      If RS2.BOF And RS2.RecordCount > 0 Then '== moved off the end so go back
        RS2.MoveFirst
      End If
    
    Case 2                                  '== move next
      If Not RS2.EOF Then RS2.MoveNext
      If RS2.EOF And RS2.RecordCount > 0 Then '==moved off the end so go back
        RS2.MoveLast
      End If
    
    Case 3                                  '== move last
      If Not RS2.EOF Then RS2.MoveLast
  
  End Select
  
  Populate_Form
  
End Sub

Private Sub Form_Activate()

  MDI_Main.mFindComp.Enabled = False  '== zavisno od toga koja je forma aktivna
  MDI_Main.mFindRec.Enabled = True    '== (komponente ili zapisi)

End Sub

Private Sub Form_Load()

  Dim ComBtn As CommandButton
  
  For Each ComBtn In cmdNavi
    SetWindowLong ComBtn.hwnd, GWL_EXSTYLE, 131076
    SetWindowPos ComBtn.hwnd, 0, 0, 0, 0, 0, wFLAGS
  Next
  
  SetWindowLong cmdClose.hwnd, GWL_EXSTYLE, 131076
  SetWindowPos cmdClose.hwnd, 0, 0, 0, 0, 0, wFLAGS
  
  MLSLoadLanguage Me
  
  RecOpen = True
  Me.Left = 0
  Me.TOp = 0
  Me.Width = 6015
  Me.Height = 6810
  
  sQry = "SELECT * FROM tbl_Work;"
  
  Conn2_Open
  
  Populate_Form
  
End Sub

Private Sub Form_Resize()

If Me.Width > 6015 Then

  Select Case Me.WindowState
  
    Case vbMaximized, vbNormal
    
    With Me
      
    lblData(0).Width = .Width - 3780
    lblData(1).Width = .Width - 3780
    lblData(2).Width = .Width - 3780
    lblData(3).Width = .Width - 1500
    lblData(4).Width = .Width - 1500
    lblData(5).Width = .Width - 1500
    
    lblTrue.Left = .Width - 795
    lblFalse.Left = .Width - 795
    
    lblInfo(6).Left = .Width - 1995
    
    cmdClose.Left = .Width - 1515
    cmdClose.TOp = .Height - 870
    
    cmdNavi(0).TOp = .Height - 870
    cmdNavi(1).TOp = .Height - 870
    cmdNavi(2).TOp = .Height - 870
    cmdNavi(3).TOp = .Height - 870
    
    ln1.X2 = .Width - 135
    ln2.X2 = .Width - 135
    
    ln1.Y1 = .Height - 1230
    ln1.Y2 = .Height - 1230
    
    ln2.Y1 = .Height - 1215
    ln2.Y2 = .Height - 1215
    
    End With
    
  End Select
  
Else

  'Me.Width = 6015
  
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)  '== Jos jednom setujemo svojstvo Enabled
  
  MDI_Main.mFindRec.Enabled = False         '== u slucaju da su obje forme ugasene
  RecOpen = False
  
  Conn2_Close
  
End Sub

Public Sub Populate_Form()

  Dim oLabel As Label
  
  If Not RS2 Is Nothing Then
    If Not RS2.EOF And Not RS2.BOF Then
      For Each oLabel In lblData
        If Not IsNull(RS2.Fields(oLabel.Tag).value) Then
          oLabel.Caption = RS2.Fields(oLabel.Tag).value
        Else
          oLabel.Caption = ""
        End If
        lblTrue.Visible = RS2.Fields("Charged").value
        lblRecCount.Caption = MLSGetString("0023") & RS2.AbsolutePosition _
        & MLSGetString("0024") & RS2.RecordCount & "]"
      Next
    End If
  End If
   
End Sub

Private Sub lblFalse_DblClick()
    
  Select Case MsgBox(MLSGetString("0025") & lblData(6).Caption & MLSGetString("0026"), vbYesNo Or vbQuestion Or vbDefaultButton2, MLSGetString("0027")) ' MLS-> "Change state of record number " ' MLS-> " from Uncharged to Charged?" ' MLS-> "Warning"

  Case vbYes
    RS2.Fields("Charged").value = True
    RS2.UpdateBatch adAffectCurrent
    lblTrue.Visible = True

  End Select
  
End Sub

