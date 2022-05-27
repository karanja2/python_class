VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   Icon            =   "frmReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sName1     As String    'path iz prve liste
Private sName2     As String    'path iz druge liste

Private Sub cmdExit_Click()

  Unload Me

End Sub

Private Sub cmdReplace_Click()

  Dim arrStr1() As String
  Dim arrStr2() As String
  Dim i1        As Integer  'broj elemenata prve liste
  Dim i2        As Integer  'broj elemenata druge liste
  Dim sTbl      As String   'tabela kojoj pristupamo
  Dim sField1   As String   'polje br1
  Dim sField3   As String   'polje br3
  Dim i         As Integer

  arrStr1 = Split(sName1, ",")
  arrStr2 = Split(sName2, ",")
  i1 = UBound(arrStr1)
  i2 = UBound(arrStr2)
  If i1 >= 0 And i2 > -1 Then
    If i1 = i2 + 1 Then
      Select Case UBound(arrStr2)
       Case 0
        sTbl = "tbl_Fasc"
        sField1 = "BrOrmana"
        sField3 = "ImeFascikle"
        sQry = "SELECT BrOrmana FROM tbl_Orm WHERE ImeOrmana = '" & arrStr2(i2) & "';"
       Case 1
        sTbl = "tbl_PodFasc"
        sField1 = "BrFascikle"
        sField3 = "ImePodfascikle"
        sQry = "SELECT BrFascikle FROM tbl_Fasc WHERE ImeFascikle = '" & arrStr2(i2) & "';"
       Case 2
        sTbl = "tbl_Doc"
        sField1 = "BrPodFasc"
        sField3 = "ImeDokumenta"
        sQry = "SELECT BrPodfasc FROM tbl_PodFasc WHERE ImePodfascikle = '" & arrStr2(i2) & "';"
      End Select
      Conn_Open
      RS.MoveFirst
      If RS.RecordCount <> 1 Then
        MsgBox "Upit je vratio nula ili više od jednog zapisa iz baze." & vbNewLine & _
       "U bazi se ne nalazi željeni podatak ili se nalaze dva ili više podataka.", vbCritical, "Greska"
        Exit Sub
      End If
      i = RS(0)
      Conn_Close
      sQry = "UPDATE " & sTbl & " SET " & sField1 & " = " & i & " WHERE " & sField3 & " = (SELECT " & sField3 & " FROM " & sTbl & " WHERE " & sField3 & " = '" & arrStr1(i1) & "');"
      Conn_Open
      Conn_Close
      MsgBox "Podatak je premješten.", vbInformation, "Arhiva"
      loadData
      sName1 = vbNullString
      sName2 = vbNullString
     Else
      MsgBox "Nedozvoljena operacija.", vbCritical, "Arhiva"
    End If
   Else
    MsgBox "Nije oznaèen podatak za premještanje.", vbInformation, "Arhiva"
  End If

End Sub

Private Sub Form_Load()

  sName1 = vbNullString
  sName2 = vbNullString
  loadData

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Unload Me

End Sub

Private Function getData(sQry As String) As ADODB.Recordset

  Set RS = New ADODB.Recordset
  RS.CursorLocation = adUseClient
  RS.Open sQry, Conn
  Set getData = RS

End Function

Private Function GetCode() As Variant

  sQry = "SELECT tbl_Main.Code_No, tbl_Main.Code, tbl_Group.Prod_Desc FROM tbl_Main, tbl_Group WHERE tbl_Group.Prod_No = tbl_Main.Prod_No;"
  Conn_Open
  Set RS = getData(sQry)
  GetCode = RS.GetRows()
  Conn_Close

End Function

Private Function GetGroup() As Variant

  sQry = "SELECT * FROM tbl_Group"
  Conn_Open
  Set RS = getData(sQry)
  GetGroup = RS.GetRows()
  Conn_Close

End Function

Private Sub loadData()

  Dim arrData As Variant
  Dim X       As Integer
  Dim Y       As Integer
  Dim nodx    As Node
  
  TreeView(0).Nodes.Clear
  TreeView(1).Nodes.Clear
  arrData = GetGroup
  
  For X = 0 To UBound(arrData, 2)
    Set nodx = TreeView(0).Nodes.Add(, tvwChild, arrData(1, X), arrData(1, X), 1)
    Set nodx = TreeView(1).Nodes.Add(, tvwChild, arrData(1, X), arrData(1, X), 1)
  Next X
  
  arrData = GetCode
  
  For Y = 0 To UBound(arrData, 2)
    Set nodx = TreeView(1).Nodes.Add(arrData(2, Y), tvwChild, arrData(1, Y), arrData(1, Y), 2, 3)
    Set nodx = TreeView(0).Nodes.Add(arrData(2, Y), tvwChild, arrData(1, Y), arrData(1, Y), 2, 3)
  Next Y
  
End Sub

Private Sub TreeView_Click(Index As Integer)

  Select Case Index
   Case 0
    sName1 = TreeView(0).SelectedItem.FullPath
   Case 1
    sName2 = TreeView(1).SelectedItem.FullPath
  End Select

End Sub


