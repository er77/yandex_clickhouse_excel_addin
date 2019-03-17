Attribute VB_Name = "mSQLConnectSVC"
Option Explicit
 
Public Const cYCHConnStringShort = "http://#YCH_Host#:#YCH_Port#"
'Public Const cYCHConnString = "http://#YCH_Host#:#YCH_Port#/?user=#YCH_user#&password=#YCH_password#&database=#YCH_DB#"
Public Const cYCHConnString = "http://#YCH_user#:#YCH_password#@#YCH_Host#:#YCH_Port#/?database=#YCH_DB#"
Public Const cYCHConnTestString = "http://#YCH_user#:#YCH_password#@#YCH_Host#:#YCH_Port#/"
Public Const cYCHConnShortString = "http://#YCH_Host#:#YCH_Port#/?database=#YCH_DB#"
' user = user&; Password = Password


Public Function setYCHConnString(ByVal vYCHCurrLogin As String, ByVal vYCHCurrPassword As String, ByVal vYCHCurrHost As String, ByVal vYCHCurrPort As String, ByVal vYCHCurrDatabase As String) As String
    
  Dim vConnString
  
     vConnString = cYCHConnString
     vConnString = Replace(vConnString, "#YCH_Host#", vYCHCurrHost)
     vConnString = Replace(vConnString, "#YCH_Port#", vYCHCurrPort)
     vConnString = Replace(vConnString, "#YCH_DB#", vYCHCurrDatabase)
     vConnString = Replace(vConnString, "#YCH_user#", vYCHCurrLogin)
     vConnString = Replace(vConnString, "#YCH_password#", vYCHCurrPassword)
     setYCHConnString = vConnString
     
End Function

Public Function setYCHConnTestString(ByVal vYCHCurrLogin As String, ByVal vYCHCurrPassword As String, ByVal vYCHCurrHost As String, ByVal vYCHCurrPort As String, ByVal vYCHCurrDatabase As String) As String
    
  Dim vConnString
  
     vConnString = cYCHConnTestString
     vConnString = Replace(vConnString, "#YCH_Host#", vYCHCurrHost)
     vConnString = Replace(vConnString, "#YCH_Port#", vYCHCurrPort)
     vConnString = Replace(vConnString, "#YCH_user#", vYCHCurrLogin)
     vConnString = Replace(vConnString, "#YCH_password#", vYCHCurrPassword)
     setYCHConnTestString = vConnString
     
End Function

Public Function setYCHConnPingString(ByVal vYCHCurrHost As String, ByVal vYCHCurrPort As String) As String
    
  Dim vConnString
  
     vConnString = UCase(cYCHConnStringShort)
     vConnString = Replace(vConnString, UCase("#YCH_Host#"), vYCHCurrHost)
     vConnString = Replace(vConnString, UCase("#YCH_Port#"), vYCHCurrPort)
     
     setYCHConnPingString = vConnString
     
End Function
 
Public Function setYCHConnShortString(ByVal vYCHCurrHost As String, ByVal vYCHCurrPort As String, ByVal vYCHCurrDatabase As String) As String
    
  Dim vConnString
  
     vConnString = cYCHConnShortString
     vConnString = Replace(vConnString, "#YCH_Host#", vYCHCurrHost)
     vConnString = Replace(vConnString, "#YCH_Port#", vYCHCurrPort)
     vConnString = Replace(vConnString, "#YCH_DB#", vYCHCurrDatabase)
     setYCHConnShortString = vConnString
      
End Function
Public Function setConnectURL(ByVal vCurrConnectLineStr As String) As String
   On Error GoTo ErrorHandler
  Dim vYCHCurrLogin, vYCHCurrPassword, vYCHCurrHost, vYCHCurrPort, vYCHCurrDatabase
  Dim vCurrLine() As String
  
  vCurrLine = Split(vCurrConnectLineStr, "|")
  

 ' Me.txtSQLCOnnNAme.Text & "'YACH'" & Me.txtYCHServer.Text & "'" & Me.txtYCHPort.Text & "'" & Me.txtYCHDatabase.Text & "'|" & Me.txtSQLUser.Text & "|" & f_XOREncryption(Me.txtSQLPass.Text, f_XOREncryption(vCurrPasswordLine, f_XOREncryption(VBA.Environ("Computername"), VBA.Environ("Username"))))
  
         vYCHCurrLogin = vCurrLine(1)
         vYCHCurrPassword = f_XORDecryption(vCurrLine(2), f_XOREncryption(vCurrPasswordLine, f_XOREncryption(VBA.Environ("Computername"), VBA.Environ("Username"))))
         
         vCurrLine() = Split(vCurrConnectLineStr, "'")
         vYCHCurrHost = vCurrLine(2)
         vYCHCurrPort = vCurrLine(3)
         vYCHCurrDatabase = vCurrLine(4)
         
 setConnectURL = setYCHConnTestString(vYCHCurrLogin, vYCHCurrPassword, vYCHCurrHost, vYCHCurrPort, vYCHCurrDatabase)
   
l_exit:
    Exit Function
ErrorHandler:
Call p_ErrorHandler(0, " Create connection filed on f_GetSQLCurrConnectString")
End Function

'setYCHConnString(ByVal vYCHCurrLogin As String, ByVal vYCHCurrPassword As String, ByVal vYCHCurrHost As String, ByVal vYCHCurrPort As String, ByVal vYCHCurrDatabase As String) As String

Function getURLfromMenuID(ByVal vMenuId As Variant)
  Dim vArr
    vArr = Split(fArrQuickConnections(vMenuId, 0), "'")
  getURLfromMenuID = setYCHConnString(fArrQuickConnections(vMenuId, 1), fArrQuickConnections(vMenuId, 2), vArr(2), vArr(3), vArr(4))
End Function

Function getURLTestfromMenuID(ByVal vMenuId As Variant)
  Dim vArr
    vArr = Split(fArrQuickConnections(vMenuId, 0), "'")
  getURLTestfromMenuID = setYCHConnTestString(fArrQuickConnections(vMenuId, 1), fArrQuickConnections(vMenuId, 2), vArr(2), vArr(3), vArr(4))
End Function

Function getURLShortfromMenuID(ByVal vMenuId As Variant)
  Dim vArr
    vArr = Split(fArrQuickConnections(vMenuId, 0), "'")
  getURLShortfromMenuID = setYCHConnShortString(vArr(2), vArr(3), vArr(4))
End Function
 

Function getMenuIDfromLink(ByVal vCurrConnectString As String)
 Dim i
 getMenuIDfromLink = -1
 
  Call ReReadConnections
  
  For i = 0 To UBound(fArrQuickConnections)
    If (InStr(UCase(fArrQuickConnections(i, 0)), UCase(vCurrConnectString)) > 0) Then
     getMenuIDfromLink = i
     Exit Function
    End If
  Next
  
   If getMenuIDfromLink < 0 Then
     MsgBox "Could not find any Quick Connnect option for parameter " & vCurrConnectString & " Please reconnect sheet", vbExclamation
     End
  End If
End Function



Function getURLFromSheetWithID(vCurrSheet As Object)

 If vCurrSheet Is Nothing Then
     MsgBox "Active sheet is not determinated ", vbExclamation
     End
 End If
  vCurrSheet.Cells(1, 1).Select
  
  If Not isTextBoxPresent("SQLConnectQ") Then
     MsgBox " Plese make connections from Quick Connect Menu ", vbExclamation
     End
 End If
  Dim vMenuId
    vMenuId = getMenuIDfromLink(getTextBoxValue("SQLConnectQ"))
    
 getURLFromSheetWithID = vMenuId & "##" & getURLShortfromMenuID(vMenuId)  'getURLfromMenuID(vMenuId)
  
End Function

Sub p_ConnectSQLByMenuId(ByVal vMenuId As Variant, Optional ByVal iSshowMsgBox As Boolean = True)

  On Error GoTo ErrorHandler
  
     If Not getYACHStatus(getURLfromMenuID(vMenuId)) Then
       MsgBox "Create connection filed", vbExclamation
      Exit Sub
     End If
     
     Call p_CreateTextBox("SQLConnectQ", "" & fArrQuickConnections(vMenuId, 0))
     MsgBox " Connected  to " & Replace(fArrQuickConnections(vMenuId, 0), "'", " ") & " by " & fArrQuickConnections(vMenuId, 1)
 
  p_RefreshRibbonNow
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " Create connection filed on p_ConnectSQLByMenuId")
    
End Sub

Sub p_TestConnectSQLByMenuId(ByVal vMenuId As Variant, Optional ByVal iSshowMsgBox As Boolean = True)

  On Error GoTo ErrorHandler
  
     If Not getYACHStatus(getURLTestfromMenuID(vMenuId)) Then
       MsgBox "Create connection filed", vbExclamation
      Exit Sub
     End If
     
     Call p_CreateTextBox("SQLConnectQ", "" & fArrQuickConnections(vMenuId, 0))
     MsgBox " Connected  to " & Replace(fArrQuickConnections(vMenuId, 0), "'", " ") & " by " & fArrQuickConnections(vMenuId, 1)
 
  p_RefreshRibbonNow
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " Create connection filed on p_ConnectSQLByMenuId")
    
End Sub
