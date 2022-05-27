Attribute VB_Name = "basMain"
Option Explicit

Public Const mcstr_DSN_Start      As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Private Const mcstr_DSN_End        As String = ";Persist Security Info=False"
Public Const mcstrDbPath          As String = "\DB\db.mdb"

Private sConnString                As String
Public sQry                        As String

Public oFSO                        As Scripting.FileSystemObject

Private Conn                       As ADODB.Connection
Public RS                          As ADODB.Recordset

Private Conn2                      As ADODB.Connection
Public RS2                         As ADODB.Recordset

Public flag                        As Boolean   '== Koristimo da naznacimo dali unosimo ProductName ili Code
Public CompOpen                    As Boolean   '== Provjera dali je frmComp otvoren
Public RecOpen                     As Boolean   '== Provjera dali je frmRecords otvoren

'== Varijable za ini fajl =='
Public sWorker                     As String  '== User
Public iAutoComponents             As Integer '== Auto start Components window
Public iAutoRecords                As Integer '== Auto start Records window
Public iUncharged                  As Integer '== Show uncharged records on start
Public iAutoStart                  As Integer '== Autostart application
Public iSysTray                    As Integer '== Minimize to sys tray
Public iSplash                     As Integer '== Show splash screen
Public sLaguage                    As String  '== Default language
Public iCompress                   As Integer '== compress database on exit
Public iBackup                     As Integer '== Backup database on exit
Public sBackupPath                 As String  '== Path for backup

Public Sub Conn_Close()
  On Error Resume Next
  
  RS.Close
  Set RS = Nothing
  Conn.Close
  Set Conn = Nothing
  sQry = vbNullString

End Sub

Public Sub Conn_Open()
  On Error GoTo greska
  
  Set Conn = New ADODB.Connection
  Conn.Open sConnString
  Set RS = New ADODB.Recordset
  RS.CursorLocation = adUseClient
  RS.Open sQry, Conn, adOpenStatic, adLockBatchOptimistic ' adLockReadOnly

Exit Sub

greska:
  Error_Handler

End Sub

Public Sub Conn2_Close()
  On Error Resume Next
  
  RS2.Close
  Set RS = Nothing
  Conn2.Close
  Set Conn2 = Nothing
  sQry = vbNullString

End Sub

Public Sub Conn2_Open()
  On Error GoTo greska
  
  Set Conn2 = New ADODB.Connection
  Conn2.Open sConnString
  Set RS2 = New ADODB.Recordset
  RS2.CursorLocation = adUseClient
  RS2.Open sQry, Conn2, adOpenStatic, adLockBatchOptimistic ' adLockReadOnly

Exit Sub

greska:
  Error_Handler

End Sub

Public Sub Error_Handler()

  Select Case Err.Number
   Case -2147467259
    Conn_Close
    MsgBox MLSGetString("0038") _
            & vbCrLf & MLSGetString("0039") _
            , vbCritical, MLSGetString("0040")
   Case Else
    MsgBox MLSGetString("0041") & Err.Number & vbNewLine & _
       MLSGetString("0042") & Err.Description, vbInformation, MLSGetString("0043") ' MLS-> "New Error!!!"
  End Select

End Sub

Public Sub Main()

  Set oFSO = New Scripting.FileSystemObject
  
  If oFSO.FileExists(oFSO.BuildPath(App.Path, mcstrDbPath)) Then
  
    sConnString = mcstr_DSN_Start & oFSO.BuildPath(App.Path, mcstrDbPath) & mcstr_DSN_End
    
    iSplash = GetProfile("Splash", "Value", "0", App.Path & "\service.ini")
    
    If iSplash = 1 Then
      frmSplash.Show
    End If
    
    iUncharged = GetProfile("ShowUncharged", "Value", "0", App.Path & "\service.ini")
    
    sLaguage = GetProfile("Language", "Value", "0", App.Path & "\service.ini")
    
    MDI_Main.Show
    
   Else
   
    MsgBox MLSGetString("0044"), vbCritical, MLSGetString("0045")

    
  End If
  
  Set oFSO = Nothing
  
End Sub
