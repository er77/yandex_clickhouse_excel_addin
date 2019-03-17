Attribute VB_Name = "mFileQuickConnSVC"
Option Explicit
Option Compare Text

Public vRibbonSetFileName As String
Public fArrQuickConnections() As String

Sub ReReadConnections()

  If (Not fArrQuickConnections) = -1 Then
    Call p_ReadConnections
  End If

End Sub

Sub SaveNewPasswordLine(vConnStringForDeleting As String, vNewConnString As String)

Call DeleteLineFromCfg(vConnStringForDeleting)

    Call GetRibonConnectionFileName
       SetAttr vRibbonSetFileName, vbNormal
        
        Open vRibbonSetFileName For Append As #1
         Print #1, vNewConnString
        Close #1
        
       SetAttr vRibbonSetFileName, vbHidden
       Call p_ReadConnections
       
       MsgBox vRibbonSetFileName & " was updated"
       
End Sub
 
 
Sub GetRibonConnectionFileName()
 On Error Resume Next
 
Dim objFolders As Object
Set objFolders = CreateObject("WScript.Shell").SpecialFolders
 
    vRibbonSetFileName = objFolders("mydocuments") & "\yach_ribon.cfg"
 

' MsgBox vRibbonSetFileName
  If (Dir(vRibbonSetFileName) = "") Then
    Open vRibbonSetFileName For Output As #1
        Write #1, ""
        Close #1
  End If
 SetAttr vRibbonSetFileName, vbNormal
 
 Set objFolders = Nothing
 
 If Err.Number <> 0 Then
   Err.Clear
 End If
 
End Sub
 


Public Sub DeleteLineFromCfg(vDeletedString As String)
 On Error GoTo ErrorHandler

    Dim vArrOfStrings() As String, vCurrStr As String
    Dim i As Long, J As Long
    Dim vCurrArrayLine() As String
    
    vDeletedString = UCase(vDeletedString)
    
    Call GetRibonConnectionFileName
      SetAttr vRibbonSetFileName, vbNormal
    Open vRibbonSetFileName For Input As 1
    i = 0
    J = 0
     DoEvents
    Do Until EOF(1)
        J = J + 1
        Line Input #1, vCurrStr
          vCurrArrayLine() = Split(vCurrStr, "|")
           If (UBound(vCurrArrayLine) > 1) Then
             If (InStr(UCase(vCurrStr), vDeletedString) = 0) Then
                i = i + 1
                ReDim Preserve vArrOfStrings(1 To i)
                vArrOfStrings(i) = vCurrStr
             End If
            End If
    Loop
    Close #1
    J = i
    
    'Write array to file
    Open vRibbonSetFileName For Output As 1
    
    For i = 1 To J
      DoEvents
        Print #1, vArrOfStrings(i)
    Next i
    Close #1

l_exit:
    SetAttr vRibbonSetFileName, vbHidden
    Exit Sub
ErrorHandler:
  Call p_ErrorHandler(0, "DeleteLineFromCfg")
    
End Sub


Public Sub p_ReadConnections()
 On Error GoTo ErrorHandler
 
     fArrQuickConnections = f_getConnections
    
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, "p_ReadConnections")
End Sub

 
 Public Function f_getConnections() As Variant
 On Error GoTo ErrorHandler

    Dim vCurrStr As String
    Dim fCurrConnections(50, 2) As String
    Dim i, J, q
    Dim vCurrArrayLine() As String
    Dim vTempArrLine(1, 5) As String
    Call GetRibonConnectionFileName
    
    SetAttr vRibbonSetFileName, vbNormal
         
    If Dir(vRibbonSetFileName) = "" Then
      Exit Function
    End If
 
    Open vRibbonSetFileName For Input As 1
    i = 0
     DoEvents
    Do Until EOF(1)
     
        Line Input #1, vCurrStr
        vCurrArrayLine() = Split(vCurrStr, "|")
        
        If UBound(vCurrArrayLine) > 1 Then
            For J = 0 To UBound(vCurrArrayLine)
                fCurrConnections(i, J) = vCurrArrayLine(J)
            Next
        End If
       i = i + 1
    Loop
    Close #1
    
    Call DeleteLineFromCfg("SS")  ' delete trash
  
    Call SetAttr(vRibbonSetFileName, vbHidden)
    
      DoEvents
    ' buble sort
      For i = 0 To UBound(fCurrConnections)
        For J = 0 To (UBound(fCurrConnections) - 1)
         If fCurrConnections(J, 0) <> "" Then
           If fCurrConnections(J, 0) > fCurrConnections(J + 1, 0) Then
            For q = 0 To 2
                vTempArrLine(1, q) = fCurrConnections(J, q)
            Next
      
            For q = 0 To 2
                fCurrConnections(J, q) = fCurrConnections(J + 1, q)
            Next
            
            For q = 0 To 2
                fCurrConnections(J + 1, q) = vTempArrLine(1, q)
            Next
         
            End If
         End If
        Next
      Next
      
     f_getConnections = fCurrConnections
    
l_exit:
    Exit Function
ErrorHandler:
Call p_ErrorHandler(0, "p_ReadConnections")
End Function





 


