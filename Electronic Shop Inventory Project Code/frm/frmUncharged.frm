VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUncharged 
   Caption         =   "Uncharged records"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUncharged.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   5535
   Begin MSComctlLib.ImageList ImageList 
      Left            =   2460
      Top             =   1740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUncharged.frx":06EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUncharged.frx":0C84
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUncharged.frx":121E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmUncharged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  Dim l As Integer
  Dim i As Integer
  
  MLSLoadLanguage Me
  
  LV.ColumnHeaders(1).Text = MLSGetString("0069")
  LV.ColumnHeaders(2).Text = MLSGetString("0070")
  LV.ColumnHeaders(3).Text = MLSGetString("0071")
  LV.ColumnHeaders(4).Text = MLSGetString("0072")
  LV.ColumnHeaders(5).Text = MLSGetString("0073")
  LV.ColumnHeaders(6).Text = MLSGetString("0074")
  
  sQry = "SELECT * FROM tbl_Work WHERE Charged = False;"

  Conn_Open
  
  For i = 0 To RS.RecordCount - 1
  
    l = 1
    
    LV.ListItems.Add (l), , RS.Fields("No"), , 1
    LV.ListItems(l).SubItems(1) = RS.Fields("Worker")
    LV.ListItems(l).SubItems(2) = RS.Fields("Date_")
    LV.ListItems(l).SubItems(3) = RS.Fields("Price")
    LV.ListItems(l).SubItems(4) = RS.Fields("Parts_Sold")
    LV.ListItems(l).SubItems(5) = RS.Fields("Customer_Info")
    
    l = l + 1
    
    RS.MoveNext
    
  Next i

End Sub

Private Sub Form_Resize()

If Me.Height > 4485 Then
  Select Case Me.WindowState
  Case vbMaximized, vbNormal
  
    LV.Width = Me.Width - 135
    LV.Height = Me.Height - 435
  
  End Select
Else
  Me.Height = 4485
End If
End Sub
