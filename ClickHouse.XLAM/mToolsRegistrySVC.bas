Attribute VB_Name = "mToolsRegistrySVC"
Option Explicit
'[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\]
'"KeepAliveTimeout"=dword:00180000
'"ReceiveTimeout"=dword:00dbba00
'"ServerInfoTimeout"=dword:00180000

Function getRegKey(vRegKey As String) As String
Dim myWS As Object
 
  On Error Resume Next
  Set myWS = CreateObject("WScript.Shell")
  getRegKey = myWS.RegRead(vRegKey)
  Set myWS = Nothing
End Function

Sub setRegKey(vRegKey As String, _
               vValue As String, _
      Optional vType As String = "REG_DWORD")
Dim myWS As Object
  Set myWS = CreateObject("WScript.Shell")
  myWS.RegWrite vRegKey, vValue, vType
  Set myWS = Nothing
End Sub

Function RegKeyExists(vRegKey As String) As Boolean
Dim myWS As Object
 
  On Error GoTo ErrorHandler
  Set myWS = CreateObject("WScript.Shell")
  myWS.RegRead vRegKey
  RegKeyExists = True
  Set myWS = Nothing
  Exit Function
   
ErrorHandler:
  RegKeyExists = False
  Set myWS = Nothing
End Function

Sub p_setRegNetworkValues()
Dim vRegStr
vRegStr = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\"

If Not RegKeyExists(vRegStr & "KeepAliveTimeout") Then
  Call setRegKey(vRegStr & "KeepAliveTimeout", "1572864", "REG_DWORD")
End If

If Not RegKeyExists(vRegStr & "ReceiveTimeout") Then
   Call setRegKey(vRegStr & "ReceiveTimeout", "14400000", "REG_DWORD")
End If

If Not RegKeyExists(vRegStr & "ServerInfoTimeout") Then
   Call setRegKey(vRegStr & "ServerInfoTimeout", "1572864", "REG_DWORD")
End If

  
 
 
End Sub


