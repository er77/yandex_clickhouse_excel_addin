Attribute VB_Name = "mRibbonDyn"
Option Explicit
 
Private Function f_makeSqlConnectMenu() As String
  Dim vCurrXML  As String
  Dim i As Integer
  
    
     vCurrXML = ""
        On Error Resume Next
        For i = 0 To UBound(fArrQuickConnections, 1)
        DoEvents
            If Not fArrQuickConnections(i, 0) = "" Then
            
               vCurrXML = vCurrXML & "<button id=""b_qConnectSQL" & i & """ label=""" & " : " & fArrQuickConnections(i, 0) & """ " _
                & "onAction=""p_qConnectAction"" imageMso=""DatabasePermissionsMenu""   />" '
                  vCurrXML = vCurrXML & "<menuSeparator id=""b_qConnectSQL2" & i & """  />"
            End If
        Next
       
       On Error GoTo ErrorHandler
     
       vCurrXML = vCurrXML & "<menuSeparator  id=""b_QuickConnect0"" />"
       vCurrXML = vCurrXML & "<button id=""b_EditSqlConnect"" label=""Manage SQL Connections"" " _
        & "onAction=""p_qConnectAction"" imageMso=""FileStartWorkflow"" />"
        
f_makeSqlConnectMenu = vCurrXML
l_exit:
    Exit Function
ErrorHandler:
 Call p_ErrorHandler(0, "f_makeEssConnectMenu")
 
End Function
 Public Sub p_SQLQuickConnect(vIRibbonControl As IRibbonControl, ByRef vXMLMenu)
 On Error GoTo ErrorHandler
    
    Dim vCurrXML As String

     
   Call p_ReadConnections
   ActiveSheet.Cells(1, 1).Select
 vCurrXML = "<menu xmlns=""" & _
           "http://schemas.microsoft.com/office/2006/01/customui"">" & vbCrLf
    
 
     vCurrXML = vCurrXML & f_makeSqlConnectMenu
 
    
    vCurrXML = vCurrXML & "</menu>"
    
    vXMLMenu = vCurrXML

   p_RefreshRibbonNow
   
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "p_QuickConnect")
  
End Sub

 







