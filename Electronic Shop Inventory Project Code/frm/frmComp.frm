VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmComp 
   Caption         =   "Components"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   8235
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7329
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
                                      '== Podesavamo svojstvo Enabled manija Find
  MDI_Main.mFindRec.Enabled = False   '== zavisno od toga koja je forma aktivna
  MDI_Main.mFindComp.Enabled = True   '== (komponente ili zapisi)
  
End Sub

Private Sub Form_Resize()

  Dim i As Integer
  
  
  If Me.Height > 4620 Then
  
  Select Case Me.WindowState
  
    Case vbMaximized, vbNormal, vbMinimized
    
      With FlexGrid
      
        .Left = 0
        .TOp = 0
        .Width = Me.Width - 135
        .Height = Me.Height - 435
        
        For i = 0 To 7
          .ColWidth(i) = (.Width / 8) - 35
        Next i
        
      End With
      
  End Select
  
  Else
  
    'Me.Height = 4620
    
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
                                            '== Jos jednom setujemo svojstvo Enabled
  MDI_Main.mFindComp.Enabled = False        '== u slucaju da su obje forme ugasene
  CompOpen = False
  
End Sub

Private Sub Form_Load()

  MLSLoadLanguage Me

  CompOpen = True
  sQry = "SELECT tbl_Group.Prod_Desc, tbl_Main.Code, tbl_Main.Descr, tbl_Main.Price1, tbl_Main.Price2, tbl_Main.Price3, tbl_Main.Quant, tbl_Main.Box FROM tbl_Group, tbl_Main WHERE tbl_Group.Prod_No = tbl_Main.Prod_No ORDER BY Prod_Desc;"
  Conn_Open
  Fill_FlexGrid
  Paint_FG
  Conn_Close
End Sub

Private Sub Fill_FlexGrid()

  Dim i As Integer

  If RS.BOF And RS.EOF Then Exit Sub
  
  RS.MoveFirst
  
  With FlexGrid
    .Visible = False
    .Clear
    '---%<---odavde-------------------------------------------------------------------'
    .Rows = RS.RecordCount + 1
    .Cols = RS.Fields.Count     '               == popunjavamo FlexGrid               '
    .Row = 1                    '               == ovaj metod je preko 250%           '
    .Col = 0                    '               == brzi od klasicnog nacina           '
    .RowSel = .Rows - 1         '               == nije provjereno na drugim          '
    .ColSel = .Cols - 1         '               == kontrolama osim FlexGrid kontrole  '
    .Clip = RS.GetString(adClipString, -1, vbTab, vbCr, vbNullString)                 '
    .Row = 1
    '---%<---d00vde-------------------------------------------------------------------'
    .TextMatrix(0, 0) = MLSGetString("0001") ' Group
    .TextMatrix(0, 1) = MLSGetString("0002") ' Code
    .TextMatrix(0, 2) = MLSGetString("0003") ' Description
    .TextMatrix(0, 3) = MLSGetString("0004") ' Price (1-10)
    .TextMatrix(0, 4) = MLSGetString("0005") ' Price (10-100)
    .TextMatrix(0, 5) = MLSGetString("0006") ' Price (100->)
    .TextMatrix(0, 6) = MLSGetString("0007") ' Quantity
    .TextMatrix(0, 7) = MLSGetString("0008") ' Box
    .Row = 0
    
    For i = 0 To .Cols - 1
      .Col = i
      .CellFontBold = True
      .CellAlignment = 3
      .ColAlignment(i) = 3
    Next i
    
    .Visible = True
    
  End With
      
End Sub

Public Sub Paint_FG()

  Dim Icrow As Integer
  Dim Icol  As Integer

  With FlexGrid
    If .Rows > 2 Then
      For Icrow = 2 To .Rows - 1 Step 2
        .Row = Icrow
        For Icol = 0 To .Cols - 1
          .Col = Icol
          .CellBackColor = &HC0FFFF
        Next Icol
      Next Icrow
    End If
  End With

End Sub
