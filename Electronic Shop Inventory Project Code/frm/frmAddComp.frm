VERSION 5.00
Begin VB.Form frmAddComp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add new component"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddComp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAction 
      Caption         =   "Exit"
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   25
      Top             =   5760
      Width           =   1155
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Clear"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   24
      Top             =   5760
      Width           =   1155
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   23
      Top             =   5760
      Width           =   1155
   End
   Begin VB.ComboBox cboProduct 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   2715
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   6
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   20
      Top             =   4920
      Width           =   2715
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   5
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   17
      Top             =   4500
      Width           =   2715
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   4
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   15
      Top             =   4080
      Width           =   2715
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   3
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   13
      Top             =   3660
      Width           =   2715
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   2
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   10
      Top             =   3240
      Width           =   2715
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   1305
      Index           =   1
      Left            =   1440
      TabIndex        =   8
      Top             =   1860
      Width           =   2715
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   0
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1440
      Width           =   2715
   End
   Begin VB.TextBox txtProduct 
      Height          =   345
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   2
      Top             =   300
      Width           =   2715
   End
   Begin VB.Label lbl 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   4260
      TabIndex        =   21
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label lbl 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   4260
      TabIndex        =   18
      Top             =   4620
      Width           =   135
   End
   Begin VB.Label lbl 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   4260
      TabIndex        =   11
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label lbl 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   4260
      TabIndex        =   6
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label lblTotalCode 
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   5340
      Width           =   2655
   End
   Begin VB.Label lblTotalProd 
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   180
      X2              =   4380
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   180
      X2              =   4380
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label lblProduct 
      Alignment       =   1  'Right Justify
      Caption         =   "Product Name:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Box:"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   4980
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantity:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   16
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Price (<100):"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   4140
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Price (10-100):"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   12
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Price (1-10):"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   9
      Top             =   3300
      Width           =   1275
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Description:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Code:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1500
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim counter As Integer

Private Sub cmdAction_Click(Index As Integer)

  Select Case Index
  
    Case 0          '== Save
      Save_Data
      
    Case 1          '== Clear
      Clear_Data
      
    Case 2          '== Exit
      Unload Me
      
  End Select

End Sub

Private Sub Form_Load()
  
  Dim i As Integer
  Dim ComBtn As CommandButton
  
  For Each ComBtn In cmdAction
    SetWindowLong ComBtn.hwnd, GWL_EXSTYLE, 131076
    SetWindowPos ComBtn.hwnd, 0, 0, 0, 0, 0, wFLAGS
  Next
  
  MLSLoadLanguage Me
  
  sQry = "SELECT Prod_Desc FROM tbl_Group;"
  
  Conn_Open
  counter = RS.RecordCount
  lblTotalProd.Caption = MLSGetString("0046") & counter & MLSGetString("0047")

  If flag Then          'flag = TRUE, znaci da unosimo novu kategoriju. Svi elementi na formi
                        '             su ugaseni osim tekst boksa Product Name
    For i = 0 To 6
      txtInput(i).Enabled = False          '== Gasimo tekst polja
      txtInput(i).BackColor = &H8000000F   '== Bojamo ih u sivo
      lblInfo(i).Enabled = False           '== Gasimo labele
    Next i
    
    For i = 0 To 3              '== Gasimo zvjezdice za obavezna polja
      lbl(i).Visible = False
    Next i
    
    cboProduct.Visible = False
    
    Conn_Close
    
  Else                  'flag = FALSE, znaci da unosimo novu komponentu. Kategoriju biramo iz kombo
                        '              boksa na vrhu forme. Sada je ugasen samo TextBox Product Name
    txtProduct.Visible = False
    
    If RS.EOF And RS.BOF Then Exit Sub
    
    RS.MoveFirst
    Do Until RS.EOF                         '== Punimo combo box
      cboProduct.AddItem (RS("Prod_Desc"))
      RS.MoveNext
    Loop
    cboProduct.Text = cboProduct.List(0)
    Conn_Close
    
    sQry = "SELECT * FROM tbl_Main" '== Ovo nam treba da provjerimo ukupan broj
    Conn_Open                       '== zapisa za info labelu na dnu.
    counter = RS.RecordCount
    lblTotalCode.Caption = MLSGetString("0048") & counter & MLSGetString("0049")
    Conn_Close
    
  End If
    
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
  txtInput(Index).SelStart = 0
  txtInput(Index).SelLength = (Len(txtInput(Index).Text))

End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)

  Select Case Index
    Case 2, 3, 4
      KeyAscii = IIf(Not KeyAscii = 8 And Not KeyAscii = 46 And Not IsNumeric(Chr$(KeyAscii)), 0, KeyAscii)
    Case 5, 6
      KeyAscii = IIf(Not KeyAscii = 8 And Not IsNumeric(Chr$(KeyAscii)), 0, KeyAscii)
  End Select
  
End Sub

Private Function IsCompOK() As Boolean

  IsCompOK = True
  
  If txtInput(0).Text = vbNullString Then
    IsCompOK = False
  ElseIf txtInput(2).Text = vbNullString Then
    IsCompOK = False
  ElseIf txtInput(5).Text = vbNullString Then
    IsCompOK = False
  ElseIf txtInput(6).Text = vbNullString Then
    IsCompOK = False
  End If
  
  sQry = "SELECT Code From tbl_Main"
  Conn_Open
  
  If RS.EOF And RS.BOF Then
    Conn_Close
    Exit Function
  End If
  
  RS.MoveFirst
  Do Until RS.EOF
    If txtInput(0).Text = (RS("Code")) Then
      IsCompOK = False
    End If
    RS.MoveNext
  Loop
  Conn_Close
  
End Function

Private Function IsCatOK() As Boolean
  
  IsCatOK = True
    
  If txtProduct.Text = vbNullString Then
    IsCatOK = False
  End If
  
  sQry = "SELECT * FROM tbl_Group"
  Conn_Open
  
  If RS.BOF And RS.EOF Then Exit Function
  
  RS.MoveFirst
  Do Until RS.EOF
    If txtProduct.Text = (RS("Prod_Desc")) Then
      IsCatOK = False
    End If
    RS.MoveNext
  Loop
  Conn_Close
  
End Function

Private Sub Save_Data()

  If flag Then    '== Save Category
  
      If IsCatOK Then
      
          sQry = "INSERT INTO tbl_Group (Prod_Desc) VALUES ('" & txtProduct.Text & "');"
          Conn_Open
          Conn_Close
          counter = counter + 1
          lblTotalProd.Caption = MLSGetString("0050") & counter & MLSGetString("0051")
          Clear_Data
        
      Else
      
          MsgBox MLSGetString("0052") _
                & vbCrLf & MLSGetString("0053") _
                , vbCritical Or vbDefaultButton1, MLSGetString("0054")
    
          txtProduct.SelStart = 0
          txtProduct.SelLength = Len(txtProduct.Text)
        
      End If
      
      txtProduct.SetFocus
    
  Else            '== Save Component
  
      If cboProduct.Text <> vbNullString Then
      
          If IsCompOK Then
          
              sQry = "INSERT INTO tbl_Main (Prod_No, Code, Descr, Price1, Price2, Price3, Quant, Box) SELECT t1.Prod_No, '" _
                    & txtInput(0).Text & "', '" & txtInput(1).Text & " ', '" & txtInput(2).Text _
                    & "', '" & txtInput(3).Text & " ', '" & txtInput(4).Text & " ', '" & txtInput(5).Text _
                    & "', '" & txtInput(6).Text & "' FROM tbl_Group AS t1 WHERE t1.Prod_Desc = '" & cboProduct.Text & "';"
              Conn_Open
              Conn_Close
              counter = counter + 1
              lblTotalCode.Caption = MLSGetString("0055") & counter & MLSGetString("0056")
              Clear_Data
            
          Else
          
              MsgBox MLSGetString("0057") _
                    & vbCrLf & MLSGetString("0058") _
                    , vbCritical Or vbDefaultButton1, MLSGetString("0059")
        
              txtInput(0).SelStart = 0
              txtInput(0).SelLength = Len(txtInput(0))
            
          End If
        
      txtInput(0).SetFocus
      
      Else
      
          MsgBox MLSGetString("0091"), vbExclamation, MLSGetString("0090")
        
      End If
    
  End If
  
End Sub

Private Sub Clear_Data()

  Dim i As Integer
  
  txtProduct.Text = vbNullString
  
  For i = 0 To 6
  
    txtInput(i).Text = vbNullString
    
  Next i

End Sub
