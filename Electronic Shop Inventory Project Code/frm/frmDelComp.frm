VERSION 5.00
Begin VB.Form frmDelComp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete component"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDelComp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3300
      TabIndex        =   5
      Top             =   4740
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Default         =   -1  'True
      Height          =   375
      Left            =   2100
      TabIndex        =   4
      Top             =   4740
      Width           =   1095
   End
   Begin VB.ListBox lstComponent 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4335
   End
   Begin VB.ComboBox cboCategory 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   60
      X2              =   4440
      Y1              =   1000
      Y2              =   1000
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   60
      X2              =   4440
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label lblInfo 
      Caption         =   "Select Componnent"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1260
      Width           =   4275
   End
   Begin VB.Label lblInfo 
      Caption         =   "Select Category"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4275
   End
End
Attribute VB_Name = "frmDelComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sListData As String     '== Sadrzi oznaceni podatak iz ListBoxa

Private Sub cboCategory_Click()

  If flag = False Then    'flag = FALSE, nista nije ugaseno. Punimo list kontrolu.
  
    sListData = vbNullString
    lstComponent.Clear
  
    sQry = "SELECT Code FROM tbl_Main WHERE Prod_No = (SELECT Prod_No FROM tbl_Group WHERE Prod_Desc = '" & cboCategory.Text & "');"
    
    Conn_Open
    
    Do Until RS.EOF
      lstComponent.AddItem (RS("Code"))
      RS.MoveNext
    Loop
    
    Conn_Close
  
  End If
    
End Sub

Private Sub cmdCancel_Click()
  
  Unload Me
  
End Sub

Private Sub cmdDelete_Click()

  If flag Then    'flag = TRUE, znaci da brisemo kategoriju.
   
  
    sQry = "SELECT * FROM tbl_Main WHERE Prod_No = (SELECT Prod_No FROM tbl_Group WHERE Prod_Desc = '" & cboCategory.Text & "');"
    
    Conn_Open
    
    If RS.RecordCount = 0 Then
    
      Conn_Close
      
      sQry = "DELETE * FROM tbl_Group WHERE Prod_Desc = '" & cboCategory.Text & "';"
      
      Conn_Open
      
      Conn_Close
      MsgBox MLSGetString("0060") & cboCategory.Text & MLSGetString("0061"), vbInformation Or vbDefaultButton1, MLSGetString("0062") ' MLS-> "Category" ' MLS-> " deleted" ' MLS-> "Information"
      cboCategory.RemoveItem (cboCategory.ListIndex)
      
      If cboCategory.Text <> vbNullString Then '== Za slucaj da je poslednja stavka u listi javice se greska raed only property
        cboCategory.Text = cboCategory.List(0)
      End If
      
    Else      'flag = FALSE, znaci da brisemo komponentu
      MsgBox MLSGetString("0063") & cboCategory.Text & MLSGetString("0064") _
            & vbCrLf & MLSGetString("0065") _
            , vbExclamation Or vbDefaultButton1, MLSGetString("0066")

    End If
    
    Conn_Close
  
  Else
  
    If sListData = vbNullString Then
      MsgBox MLSGetString("0067"), vbInformation Or vbDefaultButton1, MLSGetString("0068")
    Else
      sQry = "DELETE * FROM tbl_Main WHERE tbl_Main.Code = '" & sListData & "';"
      Conn_Open
      Conn_Close
      lstComponent.RemoveItem (lstComponent.ListIndex)
      sListData = vbNullString
    End If
    
  End If

End Sub

Private Sub Form_Load()
  
  SetWindowLong cmdDelete.hwnd, GWL_EXSTYLE, 131076
  SetWindowPos cmdDelete.hwnd, 0, 0, 0, 0, 0, wFLAGS
  SetWindowLong cmdCancel.hwnd, GWL_EXSTYLE, 131076
  SetWindowPos cmdCancel.hwnd, 0, 0, 0, 0, 0, wFLAGS
  
  MLSLoadLanguage Me
  
  sQry = "SELECT * FROM tbl_Group"
  Conn_Open
  If RS.BOF And RS.EOF Then Exit Sub
  RS.MoveFirst
  
  Do Until RS.EOF                         '== Punimo combo box
    cboCategory.AddItem (RS("Prod_Desc"))
    RS.MoveNext
  Loop
  
  Conn_Close
  cboCategory.Text = cboCategory.List(0)

  If flag Then    'flag = TRUE, znaci da brisemo kategoriju. List kontrola je ugasena
    
    lstComponent.Enabled = False
    lstComponent.BackColor = &H8000000F
    lblInfo(1).Enabled = False

  End If
    
End Sub

Private Sub lstComponent_Click()

  sListData = lstComponent.Text
  
End Sub
