Attribute VB_Name = "mRibbonAct"
Option Explicit
 Public vLastSheetName As String
  
 
 Public Sub p_qConnectAction(vIRibbonControl As IRibbonControl)
 
  On Error GoTo ErrorHandler
  Dim vtFriendlyName As String
  Dim isConnection
  
  isConnection = True
    If vIRibbonControl.ID = "b_EditSqlConnect" Then
      fManageYCHLinks.Show (1)
      isConnection = False
    Else
       Call p_TestConnectSQLByMenuId(Replace(vIRibbonControl.ID, "b_qConnectSQL", ""))
    End If
     
    
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, "p_qConnectAction")
 

End Sub
 

Sub p_SQLAbout(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler

    MsgBox "Essbase Act! Excel Ribbon  v.07.00 " & vbNewLine & " It is free under BSD license " _
          & vbNewLine & " developer: er@essbase.ru "
   
  
     p_RefreshRibbonNow
 
l_exit:
    Exit Sub
ErrorHandler:
  Call p_ErrorHandler(0, "p_About")
End Sub


Sub p_SQLSheetInfo(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
  Range(getCurrentTBRangeName).Select
  
 Application.CommandBars.ExecuteMso ("DataRangeProperties")
 If Err.Number <> 0 Then
   Err.Clear
   p_RefreshRibbonNow
 End If
l_exit:
    Exit Sub
ErrorHandler:
  Call p_ErrorHandler(0, "p_SQLSheetInfo")
End Sub

 

Sub p_ExcelCreatePivot(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
Call p_setExcelCalcOff
   Range(getCurrentTBRangeName).Select
  Application.CommandBars.ExecuteMso ("PivotTableSuggestion")
 Call p_setExcelCalcOn
End Sub

Function CheckIfSheetExists(SheetName As String) As Boolean
      CheckIfSheetExists = False
    Dim ws As Worksheet
      For Each ws In Worksheets
        If SheetName = ws.Name Then
          CheckIfSheetExists = True
          Exit Function
        End If
      Next ws
End Function
                      
Sub p_SQLBackOutl(ByVal vIRibbonControl As IRibbonControl)

     Dim vCurrSheetName
       vCurrSheetName = ActiveSheet.Name
       
      If (InStr(UCase(vCurrSheetName), "OTL") > 0) Then
        If CheckIfSheetExists(vLastSheetName) Then
          On Error Resume Next
            Worksheets(vLastSheetName).Activate
             Err.Clear
             vLastSheetName = ""
          Exit Sub
        End If
         Exit Sub
       End If
       
        vLastSheetName = vCurrSheetName
         If CheckIfSheetExists("OTL") Then
          On Error Resume Next
             Worksheets("OTL").Activate
          Err.Clear
          Exit Sub
         'Else
         '  MsgBox "This button will activate  ""OTL"" page"
         End If
         
      
       
End Sub

Sub p_SQLFreezePanes(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
 ActiveWindow.FreezePanes = Not ActiveWindow.FreezePanes
 
End Sub


Sub p_SQLAutoFilter(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next

  Range(getCurrentTBRangeName).Select
  Selection.AutoFilter
 
End Sub
