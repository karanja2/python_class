VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   3120
      Width           =   1275
   End
   Begin TabDlg.SSTab SetTab 
      Height          =   2955
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   5212
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Options"
      TabPicture(0)   =   "frmSettings.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "frmSettings.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Utility"
      TabPicture(2)   =   "frmSettings.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   1755
         Index           =   2
         Left            =   -74820
         TabIndex        =   11
         Top             =   480
         Width           =   4275
         Begin VB.CommandButton cmdPath 
            Height          =   315
            Left            =   3840
            Picture         =   "frmSettings.frx":0060
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1200
            Width           =   315
         End
         Begin VB.TextBox txtBckupPath 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   180
            TabIndex        =   14
            Text            =   "C:\DbBackup.mdb"
            Top             =   1200
            Width           =   3615
         End
         Begin VB.CheckBox chkCompress 
            Appearance      =   0  'Flat
            Caption         =   "Auto compress database on exit"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   12
            Top             =   300
            Width           =   3915
         End
         Begin VB.CheckBox chkBackup 
            Appearance      =   0  'Flat
            Caption         =   "Auto backup database on exit"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   180
            TabIndex        =   13
            Top             =   780
            Width           =   3915
         End
         Begin VB.Line Line 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   0
            X2              =   4260
            Y1              =   680
            Y2              =   680
         End
         Begin VB.Line Line 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   0
            X2              =   4260
            Y1              =   660
            Y2              =   660
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2115
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   480
         Width           =   4275
         Begin VB.ComboBox cboLanguage 
            Height          =   315
            ItemData        =   "frmSettings.frx":05EA
            Left            =   1860
            List            =   "frmSettings.frx":05F1
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1500
            Width           =   2115
         End
         Begin VB.CheckBox chkSplash 
            Appearance      =   0  'Flat
            Caption         =   "Show splash screen on startup"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   8
            Top             =   1020
            Width           =   3915
         End
         Begin VB.CheckBox chgSysTray 
            Appearance      =   0  'Flat
            Caption         =   "Minimize program to system tray"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   7
            Top             =   660
            Width           =   3915
         End
         Begin VB.CheckBox chkAutoStart 
            Appearance      =   0  'Flat
            Caption         =   "Start application when Windows start"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   6
            Top             =   300
            Width           =   3915
         End
         Begin VB.Label lblLang 
            Caption         =   "Default Language:"
            Height          =   255
            Left            =   180
            TabIndex        =   10
            Top             =   1560
            Width           =   1635
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1515
         Index           =   0
         Left            =   -74820
         TabIndex        =   1
         Top             =   480
         Width           =   4275
         Begin VB.CheckBox chkWarning 
            Appearance      =   0  'Flat
            Caption         =   "Show uncharged records when application start"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   4
            Top             =   1020
            Width           =   3735
         End
         Begin VB.CheckBox chkComponents 
            Appearance      =   0  'Flat
            Caption         =   "Open components window when application start"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   2
            Top             =   300
            Width           =   3915
         End
         Begin VB.CheckBox chkRecords 
            Appearance      =   0  'Flat
            Caption         =   "Open records window when application start"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   180
            TabIndex        =   3
            Top             =   660
            Width           =   3915
         End
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkBackup_Click()

  txtBckupPath.Enabled = chkBackup.value
  cmdPath.Enabled = chkBackup.value
  
End Sub


Private Sub cmdOk_Click()
  
  WriteProfile "ComponnentsWindow", "Value", chkComponents.value, App.Path & "\service.ini"
  WriteProfile "RecordsWindow", "Value", chkRecords.value, App.Path & "\service.ini"
  WriteProfile "ShowUncharged", "Value", chkWarning.value, App.Path & "\service.ini"
  WriteProfile "AutoStart", "Value", chkAutoStart.value, App.Path & "\service.ini"
  WriteProfile "SysTray", "Value", chgSysTray.value, App.Path & "\service.ini"
  WriteProfile "Splash", "Value", chkSplash.value, App.Path & "\service.ini"
  WriteProfile "Language", "Value", cboLanguage.Text, App.Path & "\service.ini"
  WriteProfile "AutoCompressDB", "Value", chkCompress.value, App.Path & "\service.ini"
  WriteProfile "AudoBackupDB", "Value", chkBackup.value, App.Path & "\service.ini"
  WriteProfile "BackupPath", "Value", txtBckupPath.Text, App.Path & "\service.ini"
  
  iAutoComponents = chkComponents.value
  iAutoRecords = chkRecords.value
  iUncharged = chkWarning.value
  iAutoStart = chkAutoStart.value
  iSysTray = chgSysTray.value
  iSplash = chkSplash.value
  sLaguage = cboLanguage.Text
  iCompress = chkCompress.value
  iBackup = chkBackup.value
  sBackupPath = txtBckupPath.Text
  
  If iAutoStart = 1 Then
    
    SetRegValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title, App.Path & "\" & App.Title
  
  Else
    
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title

  End If
    
  Unload Me
  
End Sub

Private Sub cmdPath_Click()

  Dim sFileName As String
  
  VBGetSaveFileName sFileName, MLSGetString("0028"), , MLSGetString("0029"), , , MLSGetString("0030"), MLSGetString("0031") ' MLS-> "DbBackup.mdb" ' MLS-> "Database File (*.mdb)|*.mdb" ' MLS-> "Database Backup" ' MLS-> "mdb"
  txtBckupPath.Text = sFileName
  
End Sub

Private Sub Form_Load()

  Dim i As Integer
  Dim sLngFileName As String

  SetWindowLong cmdOk.hwnd, GWL_EXSTYLE, 131076
  SetWindowPos cmdOk.hwnd, 0, 0, 0, 0, 0, wFLAGS
  SetWindowLong cmdPath.hwnd, GWL_EXSTYLE, 131076
  SetWindowPos cmdPath.hwnd, 0, 0, 0, 0, 0, wFLAGS
  
  MLSLoadLanguage Me

  sLngFileName = Dir(App.Path & "\*.lng")
    
  If sLngFileName <> "" Then
  
    cboLanguage.Clear
    
    i = 0
    Do While sLngFileName <> ""
      cboLanguage.List(i) = Mid(sLngFileName, 1, Len(sLngFileName) - 4)
      sLngFileName = Dir
      i = i + 1
    Loop

    sLaguage = MLSReadINI(App.Path & "\" & "LangSetting.ini", "Language", "CurrentLanguage")
    
  End If

  SetTab.Tab = 0
 
  chkComponents.value = GetProfile("ComponnentsWindow", "Value", "0", App.Path & "\service.ini")
  chkRecords.value = GetProfile("RecordsWindow", "Value", "0", App.Path & "\service.ini")
  chkWarning.value = GetProfile("ShowUncharged", "Value", "0", App.Path & "\service.ini")
  chkAutoStart.value = GetProfile("AutoStart", "Value", "0", App.Path & "\service.ini")
  chgSysTray.value = GetProfile("SysTray", "Value", "0", App.Path & "\service.ini")
  chkSplash.value = GetProfile("Splash", "Value", "0", App.Path & "\service.ini")
  cboLanguage.Text = GetProfile("Language", "Value", "English", App.Path & "\service.ini")
  chkCompress.value = GetProfile("AutoCompressDB", "Value", "0", App.Path & "\service.ini")
  chkBackup.value = GetProfile("AudoBackupDB", "Value", "0", App.Path & "\service.ini")
  txtBckupPath.Text = GetProfile("BackupPath", "Value", "0", App.Path & "\service.ini")
  
  
  If chkBackup.value = 0 Then
    txtBckupPath.Enabled = False
    cmdPath.Enabled = False
  End If

End Sub

Private Sub txtBckupPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  txtBckupPath.ToolTipText = txtBckupPath.Text
  
End Sub
