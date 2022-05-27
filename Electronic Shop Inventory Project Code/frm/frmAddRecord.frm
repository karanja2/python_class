VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAddRecord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add new record"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddRecord.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox mskEdit 
      Height          =   315
      Left            =   4260
      TabIndex        =   3
      Top             =   180
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mm-yyyy"
      Mask            =   "##-##-####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdAction 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   2
      Left            =   4620
      TabIndex        =   17
      Top             =   5280
      Width           =   1275
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Clear"
      Height          =   375
      Index           =   1
      Left            =   3180
      TabIndex        =   16
      Top             =   5280
      Width           =   1275
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1740
      TabIndex        =   15
      Top             =   5280
      Width           =   1275
   End
   Begin VB.ComboBox cboCharged 
      Height          =   315
      ItemData        =   "frmAddRecord.frx":000C
      Left            =   4260
      List            =   "frmAddRecord.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4500
      Width           =   1635
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   1515
      Index           =   3
      Left            =   1380
      TabIndex        =   9
      Top             =   2820
      Width           =   4515
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   1515
      Index           =   2
      Left            =   1380
      TabIndex        =   7
      Top             =   1140
      Width           =   4515
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   1380
      TabIndex        =   5
      Top             =   660
      Width           =   4515
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   1380
      TabIndex        =   11
      Top             =   4500
      Width           =   1395
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "€"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2700
      TabIndex        =   12
      Top             =   4500
      Width           =   315
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   6000
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   120
      X2              =   6000
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label lblWorker 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Top             =   180
      Width           =   1635
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Charged:"
      Height          =   195
      Index           =   6
      Left            =   3120
      TabIndex        =   13
      Top             =   4560
      Width           =   1035
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer Info:"
      Height          =   495
      Index           =   5
      Left            =   60
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Description:"
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Parts Sold:"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Price:"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
      Height          =   195
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Worker:"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAction_Click(Index As Integer)

  Select Case Index
  
    Case 0    '== Save Button
      Save_Data
      
    Case 1    '== Clear Button
      Clear_Controls
      
    Case 2    '== Cancel Button
      Unload Me
      
  End Select
  
End Sub

Private Sub Form_Load()

  Dim ComBtn As CommandButton
  
  cboCharged.List(0) = MLSGetString("0075")
  cboCharged.List(1) = MLSGetString("0076")
  
  For Each ComBtn In cmdAction
    SetWindowLong ComBtn.hwnd, GWL_EXSTYLE, 131076
    SetWindowPos ComBtn.hwnd, 0, 0, 0, 0, 0, wFLAGS
  Next

  MLSLoadLanguage Me
  
  Clear_Controls
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Unload frmRecord

End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)

  If Index = 0 Then

      KeyAscii = IIf(Not KeyAscii = 8 And Not KeyAscii = 46 And Not IsNumeric(Chr$(KeyAscii)), 0, KeyAscii)
  
  End If
  
End Sub

Private Sub Clear_Controls()

  Dim i As Integer
  
  For i = 0 To 3
    txtData(i).Text = vbNullString
  Next i
  
  lblWorker.Caption = sWorker
  
  cboCharged.ListIndex = 0
  
  mskEdit.Text = Format(Date, "dd-mm-yyyy")
  mskEdit.SelLength = (Len(mskEdit.Text))
  
End Sub

Private Sub Save_Data()

  Dim bCharg As Boolean
  
  bCharg = True
  
  If cboCharged.Text = MLSGetString("0033") Then
    bCharg = False
  End If

  If mskEdit.Text <> "__-__-____" Then
  
    sQry = "INSERT INTO tbl_Work (Worker, Date_, Price, Parts_Sold, Description, Customer_Info, Charged) Values ('" _
          & lblWorker.Caption & "', '" & mskEdit.Text & "', '" & txtData(0).Text & " ', '" _
          & txtData(1).Text & " ', '" & txtData(2).Text & " ', '" & txtData(3).Text & " ', " & bCharg & ");"
    
    Conn_Open
    Conn_Close
    MsgBox MLSGetString("0034"), vbInformation, MLSGetString("0035")
    Clear_Controls
  Else
  
    MsgBox MLSGetString("0036"), vbExclamation Or vbDefaultButton1, MLSGetString("0037")

  End If

End Sub

