Attribute VB_Name = "mToolsVariablesSVC"
Option Explicit
Option Compare Text


Global Const isDebug = True
Private Const C_ILLEGAL_CHARS = " /-:;!@#$%^&*()+=,<>"
 
 
' By Chip Pearson, www.cpearson.com , chip@cpearson.com
'  http://www.cpearson.com/excel/hidden.htm
Public Function f_ReadGlobalProperty(Key As String) As Variant
On Error Resume Next
  f_ReadGlobalProperty = Null
  If svc_HiddenNameExists(Key) Then
      f_ReadGlobalProperty = svc_GetHiddenNameValue(Key)
  End If
      
End Function

Public Sub p_WriteGlobalProperty(Key As String, value As Variant)
On Error Resume Next
      Dim vResult
    vResult = svc_AddHiddenName(Key, value)
End Sub

Public Function isValidName(HiddenName As String) As Boolean
Dim C As String
Dim NameNdx As Long
Dim CharNdx As Long
If Trim(HiddenName) = vbNullString Then
    isValidName = False
    Exit Function
End If
For NameNdx = 1 To Len(HiddenName)
DoEvents
    For CharNdx = 1 To Len(C_ILLEGAL_CHARS)
        If StrComp(Mid(HiddenName, NameNdx, 1), Mid(C_ILLEGAL_CHARS, CharNdx, 1), vbBinaryCompare) = 0 Then
            isValidName = False
            Exit Function
        End If
    Next CharNdx
Next NameNdx

isValidName = True


End Function


Public Function svc_HiddenNameExists(HiddenName As String) As Boolean
Dim v As Variant
On Error Resume Next

If isValidName(HiddenName) = False Then
    svc_HiddenNameExists = False
    Exit Function
End If
v = Application.ExecuteExcel4Macro(HiddenName)
On Error GoTo 0
If IsError(v) = False Then
    svc_HiddenNameExists = True
Else
    svc_HiddenNameExists = False
End If

End Function

 
 
Public Function svc_AddHiddenName(HiddenName As String, NameValue As Variant) As Boolean
Dim v As Variant
Dim Res As Variant

If isValidName(HiddenName) Then

If svc_HiddenNameExists(HiddenName) Then
   svc_DeleteHiddenName (HiddenName)
End If
 
v = Application.ExecuteExcel4Macro("SET.NAME(" & Chr(34) & HiddenName & Chr(34) & "," & Chr(34) & NameValue & Chr(34) & ")")

If IsError(v) = True Then
    svc_AddHiddenName = False
Else
    svc_AddHiddenName = True
End If

End If
End Function

Public Sub svc_DeleteHiddenName(HiddenName As String)
On Error Resume Next
    Application.ExecuteExcel4Macro ("SET.NAME(" & Chr(34) & HiddenName & Chr(34) & ")")

End Sub

Public Function svc_GetHiddenNameValue(HiddenName As String) As Variant
Dim v As Variant

    If isValidName(HiddenName) = False Then
        svc_GetHiddenNameValue = Null
        Exit Function
    End If
    
    If svc_HiddenNameExists(HiddenName:=HiddenName) = False Then
        svc_GetHiddenNameValue = Null
        Exit Function
    End If
    
    On Error Resume Next
    v = Application.ExecuteExcel4Macro(HiddenName)
    
    On Error GoTo 0
    If IsError(v) = True Then
        svc_GetHiddenNameValue = Null
        Exit Function
    End If
    
    svc_GetHiddenNameValue = v
        

End Function
 

 
 
 


