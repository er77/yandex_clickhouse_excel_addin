VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fManageYCHLinks 
   Caption         =   "Manage QlickHouse Connections"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7530
   OleObjectBlob   =   "fManageYCHLinks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fManageYCHLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private vISEmptyConfig As Boolean
Private vTempEnvNum As Integer
Private vEssBaseLink As String
Private vSQLLink As String


 

 

Private Sub fSQLCancel_Click()
 On Error Resume Next
 
   Call p_ReadConnections
   Unload Me
 
End Sub

Private Sub fManageSQLLinks_Terminate()
  On Error Resume Next
    Call fSQLCancel_Click
 
End Sub
 
 

Private Sub fSQLDelete_Click()
 On Error GoTo ErrorHandler
Dim iCount As Integer, i As Integer
Dim vCurrArrayLine() As String
Dim vCurrStr
iCount = Me.fSQLListLinks.ListCount - 1
Dim X
For i = 0 To iCount
X = iCount - i
    If Me.fSQLListLinks.Selected(X) Then
      vCurrArrayLine() = Split(fSQLListLinks.List(X), "|")
     Call DeleteLineFromCfg(Replace(getClearString(vCurrArrayLine(0)), ":", "'"))
    End If
Next i

  Call p_ReadConnections
  Call UserForm_Initialize
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, " fSQLDelete_Click")

End Sub
 

Private Sub fSqlLoad_Click()
 On Error GoTo ErrorHandler

 
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections fLoad_Click")
 
End Sub


Private Sub fSQLEdit_Click()
On Error GoTo ErrorHandler


If Me.fSQLListLinks Is Nothing Then
  MsgBox " Please select line "
  Exit Sub
End If

Dim vSTR

Dim iCount As Integer, i As Integer, J As Integer
Dim X
iCount = Me.fSQLListLinks.ListCount - 1
For i = 0 To iCount
X = iCount - i
    If Me.fSQLListLinks.Selected(X) Then
     Dim vCurrArrayLine() As String
   
     vCurrArrayLine() = Split(fSQLListLinks.List(X), "|")
     vSTR = getClearString(vCurrArrayLine(0))
     vSTR = Replace(vSTR, ":", "'")
       For J = 0 To UBound(fArrQuickConnections, 1)
        If fArrQuickConnections(J, 0) <> "" Then
         If InStr(fArrQuickConnections(J, 0), vSTR) > 0 Then
               Call getLinkToEdit(J)
         End If
        End If
      Next J
    End If
Next i
  
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage sql Conections fSQLEdit_Click")
End Sub

Private Sub clearUserInput()
 On Error GoTo ErrorHandler
 
   Me.txtSQLCOnnNAme.Text = getClearString(Me.txtSQLCOnnNAme.Text)
   Me.txtYCHServer.Text = getClearString(Me.txtYCHServer.Text)
   Me.txtYCHPort.Text = getClearString(Me.txtYCHPort.Text)
   Me.txtYCHDatabase.Text = getClearString(Me.txtYCHDatabase.Text)
   Me.txtSQLUser.Text = getClearString(Me.txtSQLUser.Text)
   Me.txtSQLPass.Text = getClearString(Me.txtSQLPass.Text)
   
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections f_clearUserInput")
End Sub

Private Function getCurrConnectString() As String
 On Error GoTo ErrorHandler
 
         getCurrConnectString = setYCHConnString(Me.txtSQLUser.Text, Me.txtSQLPass.Text, Me.txtYCHServer.Text, Me.txtYCHPort.Text, Me.txtYCHDatabase.Text)
 
l_exit:
    Exit Function
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections getCurrConnectString")
End Function

Private Function getCurrTestConnectString() As String
 On Error GoTo ErrorHandler
 
         getCurrTestConnectString = setYCHConnTestString(Me.txtSQLUser.Text, Me.txtSQLPass.Text, Me.txtYCHServer.Text, Me.txtYCHPort.Text, Me.txtYCHDatabase.Text)
 
l_exit:
    Exit Function
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections getCurrTestConnectString")
End Function

Private Function getCurrConnectStringShort() As String
 On Error GoTo ErrorHandler
 
         getCurrConnectStringShort = setYCHConnPingString(Me.txtYCHServer.Text, Me.txtYCHPort.Text)
 
l_exit:
    Exit Function
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections getCurrConnectStringShort")
End Function

Private Function getConnectStringForCleringFile() As String
 On Error GoTo ErrorHandler
 
  getConnectStringForCleringFile = Me.txtSQLCOnnNAme.Text & "'YACH'" & Me.txtYCHServer.Text & "'" & Me.txtYCHPort.Text & "'" & Me.txtYCHDatabase.Text & "'"
 
l_exit:
    Exit Function
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections getConnectStringForStoredFile")
End Function

 
Private Function getConnectStringForSavingFile() As String
 On Error GoTo ErrorHandler
 
  getConnectStringForSavingFile = getConnectStringForCleringFile & "|" & Me.txtSQLUser.Text & "|" & f_XOREncryption(Me.txtSQLPass.Text, f_XOREncryption(vCurrPasswordLine, f_XOREncryption(VBA.Environ("Computername"), VBA.Environ("Username"))))
 
l_exit:
    Exit Function
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections getConnectStringForStoredFile")
End Function

Private Sub fSqlSave_Click()
 
  On Error GoTo ErrorHandler

    Dim vArrOfStrings() As String, vCurrStr As String
    Dim i As Long, J As Long
    Dim vConnName As String
    
    Call clearUserInput
    
     If Not getYACHStatus(getCurrTestConnectString) Then
       MsgBox "Connection Test Filed", vbExclamation
      Exit Sub
     End If
    
   Call SaveNewPasswordLine(getConnectStringForCleringFile, getConnectStringForSavingFile)
   Call updateListLinks
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage SQL Conections fSqlSave_Click")
 
End Sub
 
 
Private Sub fTestConnect_Click()
 On Error Resume Next
  
  If getYACHStatus(getCurrTestConnectString) Then
    MsgBox "Test Passed", vbInformation
  Else
    MsgBox " Test Filed with string " & getCurrConnectStringShort, vbExclamation
    MsgBox Err.Description, vbExclamation
 End If
End Sub
 
Private Sub fSqlPingConnect_Click()
 On Error Resume Next
  
 If pingYACH(getCurrConnectStringShort) Then
    MsgBox "Ping Passed", vbInformation
  Else
    MsgBox " Ping Filed with string " & getCurrConnectStringShort, vbExclamation
    MsgBox Err.Description, vbExclamation
 End If
 
 End Sub

Private Sub initDefault()
    Me.txtYCHServer.Text = "127.0.0.1"
    Me.txtYCHPort.Text = "8123"
    Me.txtSQLUser.Text = "Default"
    Me.txtYCHDatabase.Text = "Default"
    Me.txtSQLPass.Text = ""
    Me.txtSQLCOnnNAme = "YACHLocal"
End Sub
 
 
Private Sub initialaseHeader()
 On Error GoTo ErrorHandler

Dim J, vTestSQL

   vISEmptyConfig = True
   
   For J = 0 To UBound(fArrQuickConnections, 1)
        If fArrQuickConnections(J, 0) <> "" Then
          vISEmptyConfig = False
        End If
    Next J
    
If vISEmptyConfig Then
   Call initDefault
End If

l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections initialaseHeader")

End Sub

Sub getLinkToEdit(vID As Integer)
 On Error GoTo ErrorHandler
Dim vCurrArray() As String
Dim vCurrArray2() As String
 
' Me.txtSQLCOnnNAme.Text & "'YACH'" & Me.txtYCHServer.Text & "'" & Me.txtYCHPort.Text & "'" & Me.txtYCHDatabase.Text & "'|" & Me.txtSQLUser.Text & "|" & f_XOREncryption(Me.txtSQLPass.Text, f_XOREncryption(vCurrPasswordLine, f_XOREncryption(VBA.Environ("Computername"), VBA.Environ("Username"))))
 
  Dim vCurrArrayLine() As String
 vCurrArrayLine() = Split(fArrQuickConnections(vID, 0), "'")
 
 If UBound(vCurrArrayLine) > 3 Then
    Me.txtSQLCOnnNAme.Text = vCurrArrayLine(0)
    Me.txtYCHServer.Text = vCurrArrayLine(2)
    Me.txtYCHPort.Text = vCurrArrayLine(3)
    Me.txtYCHDatabase.Text = vCurrArrayLine(4)
    vCurrArrayLine() = Split(fArrQuickConnections(vID, 0), "|")
    Me.txtSQLUser.Text = fArrQuickConnections(vID, 1)
 End If
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections getLinkToEdit")
End Sub

 


 
 Private Sub updateListLinks()
 On Error GoTo ErrorHandler
 Dim i As Long, J As Long
 Dim vCurrArrayLine() As String
 Dim vCurrSpace
    fSQLListLinks.Clear
     On Error Resume Next
     Call p_ReadConnections
 
    For i = 0 To UBound(fArrQuickConnections)
        If fArrQuickConnections(i, 0) <> "" Then
        
         vCurrArrayLine() = Split(fArrQuickConnections(i, 0), "'")
         vCurrSpace = " " & vCurrArrayLine(0) & ":" & vCurrArrayLine(1)
         vCurrSpace = vCurrSpace & Space(18 - Len(vCurrSpace)) & Chr(9) & vCurrArrayLine(2) & "   " & fArrQuickConnections(i, 3)
         
         Me.fSQLListLinks.AddItem vCurrSpace & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "|" & CRC16HASH(fArrQuickConnections(i, 0))     ' fArrQuickConnections(i, 0)
       
       End If
    Next i
    
    If Err.Number <> 0 Then
     Err.Clear
    End If
    
   On Error GoTo ErrorHandler
   
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections updateListLinks")
End Sub
 
  
Private Sub txtServer_Change()

End Sub



Private Sub UserForm_Initialize()
 On Error GoTo ErrorHandler
  
   Call updateListLinks
   Call initialaseHeader
   
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "Manage Conections UserForm_Initialize")
End Sub


 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 

 ' å r @ å s s b à s å . r u
