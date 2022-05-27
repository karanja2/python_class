VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDI_Main 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   ClientHeight    =   8430
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12315
   Icon            =   "MDI_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Tbar 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   3480
      Top             =   5340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Main.frx":0ECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Main.frx":15C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Main.frx":1CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Main.frx":23B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Main.frx":2AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Main.frx":31AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Main.frx":38A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Main.frx":3FA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI_Main.frx":469A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   8130
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   7232
            MinWidth        =   7232
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   706
            MinWidth        =   706
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "03/03/09"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "13:44"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuService 
      Caption         =   "&Service"
      Begin VB.Menu mComp 
         Caption         =   "View Components"
         Shortcut        =   ^C
      End
      Begin VB.Menu mRec 
         Caption         =   "View Records"
         Shortcut        =   ^R
      End
      Begin VB.Menu crtica0 
         Caption         =   "-"
      End
      Begin VB.Menu mUncharged 
         Caption         =   "Show Uncharged Records"
      End
      Begin VB.Menu crtica8 
         Caption         =   "-"
      End
      Begin VB.Menu mPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu crtica9 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mFindComp 
         Caption         =   "Find Components"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mFindRec 
         Caption         =   "Find Records"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mAddCategory 
         Caption         =   "Add &Category"
         Shortcut        =   ^T
      End
      Begin VB.Menu mAddComponent 
         Caption         =   "Add Co&mponent"
         Shortcut        =   ^M
      End
      Begin VB.Menu crtica1 
         Caption         =   "-"
      End
      Begin VB.Menu mAddRecord 
         Caption         =   "Add &New Record"
         Shortcut        =   ^D
      End
      Begin VB.Menu crtica2 
         Caption         =   "-"
      End
      Begin VB.Menu mDelCat 
         Caption         =   "Delete Category"
      End
      Begin VB.Menu mDelComp 
         Caption         =   "Delete Component"
      End
      Begin VB.Menu crtica3 
         Caption         =   "-"
      End
      Begin VB.Menu mDelRec 
         Caption         =   "Delete Record"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mBackup 
         Caption         =   "Backup Database"
      End
      Begin VB.Menu mCompress 
         Caption         =   "Compress Database"
      End
      Begin VB.Menu crtica5 
         Caption         =   "-"
      End
      Begin VB.Menu mSettings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      Begin VB.Menu mHorizontaly 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mVertically 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu crtica6 
         Caption         =   "-"
      End
      Begin VB.Menu mCloseWin 
         Caption         =   "Close All Windows"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu crtica7 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About..."
      End
   End
   Begin VB.Menu mPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu pupRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu crtica11 
         Caption         =   "-"
      End
      Begin VB.Menu pupAbout 
         Caption         =   "About"
      End
      Begin VB.Menu pupExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MDI_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mAbout_Click()

    frmAbout.Show vbModal

End Sub

Private Sub mAddRecord_Click()
  
  frmAddRecord.Show vbModal
  frmRecord.Show
  
  If RecOpen Then
  
    Unload frmRecord
    frmRecord.Show
    
  End If

End Sub

Private Sub mAddCategory_Click()

  flag = True
  frmAddComp.Show vbModal
  
  If CompOpen Then
  
    Unload frmComp
    frmComp.Show
    
  End If
  
End Sub

Private Sub mAddComponent_Click()

  flag = False
  frmAddComp.Show vbModal
  
  If CompOpen Then
  
    Unload frmComp
    frmComp.Show
    
  End If
  
End Sub

Private Sub mBackup_Click()

  Dim sDestFile As String
  
  If VBGetSaveFileName(sDestFile, , True, MLSGetString("0009"), , , MLSGetString("0010"), MLSGetString("0011")) Then ' MLS-> "Database Files (*.mdb)|*.mdb" ' MLS-> "Database Backup" ' MLS-> "mdb"
  
    BackupDB (sDestFile)
  
  End If

End Sub

Private Sub mCloseWin_Click()

  Dim oForm As Form
  
  For Each oForm In Forms
    If oForm.Name <> MDI_Main.Name Then
      Unload oForm
    End If
  Next oForm
  
End Sub

Private Sub mCompress_Click()

  Compact_DB
  
End Sub

Private Sub MDIForm_Load()

  MLSLoadLanguage Me

  Me.Caption = MLSGetString("0012") & App.Major & "." & App.Minor & "." & App.Revision & ")" ' MLS-> "Service & Workshop Inventory Application " ' MLS-> " (Ver. "
  
  SBar.Panels(1).Text = Me.Caption
  SBar.Panels(3).Text = MLSGetString("0014") & Now & ")"
  
  Tbar.Buttons(4).ButtonMenus(1).Text = MLSGetString("0078")
  Tbar.Buttons(4).ButtonMenus(2).Text = MLSGetString("0079")
  Tbar.Buttons(4).ButtonMenus(3).Text = MLSGetString("0080")
  
  Tbar.Buttons(1).ToolTipText = MLSGetString("0081")
  Tbar.Buttons(2).ToolTipText = MLSGetString("0082")
  Tbar.Buttons(3).ToolTipText = MLSGetString("0083")
  Tbar.Buttons(4).ToolTipText = MLSGetString("0084")
  Tbar.Buttons(5).ToolTipText = MLSGetString("0085")
  Tbar.Buttons(6).ToolTipText = MLSGetString("0086")
  Tbar.Buttons(8).ToolTipText = MLSGetString("0087")
  Tbar.Buttons(9).ToolTipText = MLSGetString("0088")
  Tbar.Buttons(10).ToolTipText = MLSGetString("0089")
  
  iAutoComponents = GetProfile("ComponnentsWindow", "Value", "0", App.Path & "\service.ini")
  iAutoRecords = GetProfile("RecordsWindow", "Value", "0", App.Path & "\service.ini")
  iCompress = GetProfile("AutoCompressDB", "Value", "0", App.Path & "\service.ini")
  iBackup = GetProfile("AudoBackupDB", "Value", "0", App.Path & "\service.ini")
  sBackupPath = GetProfile("BackupPath", "Value", "0", App.Path & "\service.ini")
  iSysTray = GetProfile("SysTray", "Value", "0", App.Path & "\service.ini")

  If iAutoRecords = 1 Then
    frmRecord.Show
  End If
  
  If iAutoComponents = 1 Then
    frmComp.Show
  End If
  
  If iUncharged = 1 Then
    frmUncharged.Show
  End If

  frmLogin.Show vbModal

End Sub

Private Sub mExit_Click()

  Dim frm As Form
  
  For Each frm In Forms
  
    Unload frm
    
  Next

End Sub

Private Sub mFindComp_Click()
  
  frmFindComp.Show

End Sub

Private Sub mFindRec_Click()
  
  frmFindRecord.Show

End Sub

Private Sub mComp_Click()
  
  frmComp.Show
  frmComp.SetFocus
  
End Sub

Private Sub mRec_Click()
  
  frmRecord.Show
  frmRecord.SetFocus
  
End Sub

Private Sub mDelCat_Click()
  
  flag = True
  frmDelComp.Show vbModal

End Sub

Private Sub mDelComp_Click()
  
  flag = False
  frmDelComp.Show vbModal

End Sub

Private Sub mSettings_Click()
  
  frmSettings.Show vbModal
  
    MLSLoadLanguage Me
       
    SBar.Panels(2).Text = MLSGetString("0032") & sWorker & ")"
    SBar.Panels(3).Text = MLSGetString("0014") & Now & ")"
    Me.Caption = MLSGetString("0012") & App.Major & "." & App.Minor & "." & App.Revision & ")" ' MLS-> "Service & Workshop Inventory Application " ' MLS-> " (Ver. "
    SBar.Panels(1).Text = Me.Caption
    
    Tbar.Buttons(4).ButtonMenus(1).Text = MLSGetString("0078")
    Tbar.Buttons(4).ButtonMenus(2).Text = MLSGetString("0079")
    Tbar.Buttons(4).ButtonMenus(3).Text = MLSGetString("0080")
    
    Tbar.Buttons(1).ToolTipText = MLSGetString("0081")
    Tbar.Buttons(2).ToolTipText = MLSGetString("0082")
    Tbar.Buttons(3).ToolTipText = MLSGetString("0083")
    Tbar.Buttons(4).ToolTipText = MLSGetString("0084")
    Tbar.Buttons(5).ToolTipText = MLSGetString("0085")
    Tbar.Buttons(6).ToolTipText = MLSGetString("0086")
    Tbar.Buttons(8).ToolTipText = MLSGetString("0087")
    Tbar.Buttons(9).ToolTipText = MLSGetString("0088")
    Tbar.Buttons(10).ToolTipText = MLSGetString("0089")
  
    
    If frmComp.WindowState < 3 Then
      Unload frmComp
      frmComp.Show
    End If
    
    If frmRecord.WindowState < 3 Then
      Unload frmRecord
      frmRecord.Show
    End If

End Sub

Private Sub mUncharged_Click()

  frmUncharged.Show
  frmUncharged.SetFocus
  
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

  If iCompress = 1 Then
  
    Compact_DB
    
  End If
  
  If iBackup = 1 Then
  
    BackupDB (sBackupPath)
    
  End If
  
End Sub

Private Sub BackupDB(sDest As String)

  Dim BckupFolder As String
  
  Set oFSO = New Scripting.FileSystemObject
  BckupFolder = oFSO.GetParentFolderName(sDest)
  
  If oFSO.FolderExists(BckupFolder) Then
  
    oFSO.CopyFile App.Path & "\db\db.mdb", sDest, True
    Set oFSO = Nothing
    MsgBox MLSGetString("0015") _
            & vbCrLf & "(" & sDest & ")" _
            , vbInformation Or vbDefaultButton1, MLSGetString("0016")

  Else
    
    MsgBox MLSGetString("0017") & BckupFolder & MLSGetString("0018") & vbCrLf & _
    MLSGetString("0019"), vbCritical Or vbDefaultButton1, MLSGetString("0020")
    
  End If
  
End Sub

Private Sub Compact_DB()

  Dim MSJet       As JRO.JetEngine
  Dim sTempFile As String
    
  Unload frmRecord
  Conn_Close
  Conn2_Close
  
  Set MSJet = New JRO.JetEngine
  Set oFSO = New Scripting.FileSystemObject
  
  MousePointer = 13
  
  sTempFile = App.Path & "\DB\tmp_db.bin"
  
  MSJet.CompactDatabase mcstr_DSN_Start & oFSO.BuildPath(App.Path, mcstrDbPath), mcstr_DSN_Start & sTempFile & ";Jet OLEDB:Engine Type=5"
  
  oFSO.CopyFile sTempFile, App.Path & mcstrDbPath, True

  oFSO.DeleteFile sTempFile
  
  MousePointer = 0
  
  Set oFSO = Nothing
  Set MSJet = Nothing
  
  MsgBox MLSGetString("0021"), vbInformation Or vbDefaultButton1, MLSGetString("0022") ' MLS-> "Database compresion succesful." ' MLS-> "Information"

End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim Sys As Long

  Sys = X / Screen.TwipsPerPixelX
  
  Select Case Sys
  
   Case WM_RBUTTONDOWN
   
      Me.PopupMenu mPopup
    
   Case WM_LBUTTONDBLCLK
   
      WindowState = vbNormal
      Me.Show
    
  End Select

End Sub

Private Sub MDIForm_Resize() '== minimiziranje forme u sys tray

  If WindowState = vbMinimized Then
  
    If iSysTray = 0 Then
    
      Exit Sub
      
    End If
    
    Me.Hide
    
    With nid
    
      .cbSize = Len(nid)
      .hwnd = Me.hwnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
      .szTip = Me.Caption & vbNullChar
      
    End With
    
    Shell_NotifyIcon NIM_ADD, nid
    
   Else
   
    Shell_NotifyIcon NIM_DELETE, nid
    
  End If
  
End Sub

Private Sub pupAbout_Click()

  frmAbout.Show vbModal

End Sub

Private Sub pupExit_Click()

  mExit_Click
  
End Sub

Private Sub pupRestore_Click()

  WindowState = vbNormal
  Me.Show

End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)

  Select Case Button.Index
  
    Case 1
      frmComp.Show              '==View Components
''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case 2
      frmRecord.Show            '== View Records
''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case 3
      frmUncharged.Show         '== Show uncharged
''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case 5
      If mFindRec.Enabled Then  '== Find...
        frmFindRecord.Show
      End If
      
      If mFindComp.Enabled Then
        frmFindComp.Show
      End If
                          
    Case 6
      mSettings_Click          '== Settings
''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case 8
      frmAbout.Show vbModal     '== About
''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case 11
      mExit_Click               '== Exit
''''''''''''''''''''''''''''''''''''''''''''''''''''
  End Select
      
End Sub

Private Sub Tbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

  Select Case ButtonMenu.Index
    Case 1
      mAddCategory_Click
      
    Case 2
      mAddComponent_Click
      
    Case 3
      mAddRecord_Click
      
  End Select

End Sub
   
Private Sub mIcons_Click()
  Me.Arrange vbArrangeIcons
End Sub

Private Sub mVertically_Click()
  Me.Arrange vbTileVertical
End Sub

Private Sub mHorizontaly_Click()
  Me.Arrange vbTileHorizontal
End Sub

Private Sub mCascade_Click()
  Me.Arrange vbCascade
End Sub
