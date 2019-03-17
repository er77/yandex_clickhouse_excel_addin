Attribute VB_Name = "mSQLActions"
Option Explicit

Sub p_SQLRetrieve(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
  If f_SQLgetEnabled Then
    Call p_MoveActiveCellToHeader
    Call p_SqlQuery(ActiveSheet)
  Else
    p_RefreshRibbonNow
   End If
 Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " p_SQLRetrieve ")
End Sub

Sub p_RefreshAll(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
 Dim i
         For i = 1 To ActiveWorkbook.Worksheets.Count
              ActiveWorkbook.Worksheets(i).Select
            If f_SQLgetEnabled Then
                Call p_MoveActiveCellToHeader
                Call p_SqlQuery(ActiveSheet)
            End If
         Next i
         
 Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " p_RefreshAll ")
End Sub

Sub p_SQLExecStored(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
  If f_SQLgetEnabled Then
     Call fExecuteSQL.Show(1)
  Else
       p_RefreshRibbonNow
   End If
 Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " p_SQLExecStored ")
End Sub

Sub p_SaveSQL(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
  If f_SQLgetEnabled Then
     Call fManageYCHLinks.Show(1)
  Else
       p_RefreshRibbonNow
   End If
   
 Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " p_SaveSQL ")
End Sub

Sub p_SQLCreatePivot(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
  If f_SQLgetEnabled Then
   ' Call p_CreatePivot(ActiveSheet)
  Else
       p_RefreshRibbonNow
   End If
 Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " p_SQLCreatePivot ")
End Sub

 Function getCurrentTBRangeName() As String
On Error GoTo ErrorHandler
   getCurrentTBRangeName = Replace(fArrQuickConnections(getMenuIDfromLink(getTextBoxValue("SQLConnectQ")), 0), "'", "_")
   getCurrentTBRangeName = getCurrentTBRangeName & CRC16HASH(ActiveSheet.Name)
 Exit Function
ErrorHandler:
Call p_ErrorHandler(0, " getCurrentTBRange ")
 End Function

Sub p_CheckCell()
On Error GoTo ErrorHandler
  If Not f_SQLgetEnabled Then
    Call p_RefreshRibbonNow
    End
  End If
  
  If Intersect(ActiveCell, Range(getCurrentTBRangeName)) Is Nothing Then
     MsgBox "You need to choose on cell from the range"
       Range(getCurrentTBRangeName).Select
    End
  End If
 Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " p_CheckCell ")
  
End Sub

Sub p_MoveActiveCellToHeader()
On Error Resume Next
Dim vSTR
Dim vAddr() As String


   Range(getCurrentTBRangeName).Clear
   vSTR = Range(getCurrentTBRangeName).AddressLocal
   vAddr = Split(vSTR, ":")
   ActiveSheet.Range(vAddr(0)).Select
   
If Err.Number <> 0 Then
   ActiveSheet.Range("A1").Select
   Err.Clear
End If

  
End Sub


Function getHeaderValue()
On Error GoTo ErrorHandler
Dim vAddr() As String, vArr() As String

   getHeaderValue = Range(getCurrentTBRangeName).AddressLocal
   vAddr = Split(getHeaderValue, "$")
   vArr = Split(ActiveCell.Address(True, False), "$")
   getHeaderValue = ActiveSheet.Range(vArr(0) & Replace(vAddr(2), ":", "")).value
 Exit Function
ErrorHandler:
Call p_ErrorHandler(0, " getHeaderValue ")

End Function

Sub p_SQLKeepOnly(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
Dim vSQL

   Call p_CheckCell
   
   vSQL = "Select * from ( " & getTextBoxValue("SqlQ") & " ) bb where " & getHeaderValue
  
   If (IsNumeric(ActiveCell.value)) Then
     vSQL = vSQL & " = " & ActiveCell.value
   Else
     vSQL = vSQL & " like '%" & ActiveCell.value & "%'"
   End If
   
   Call p_MoveActiveCellToHeader
   Call p_CreateTextBox("SqlQ", "" & vSQL)
   Call p_SqlQuery(ActiveSheet)
 Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " p_SQLKeepOnly ")

End Sub

Sub p_SQLRemoveOnly(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
Dim vSQL

   Call p_CheckCell
   
   vSQL = "Select * from ( " & getTextBoxValue("SqlQ") & " ) bb where " & getHeaderValue
  
   If (IsNumeric(ActiveCell.value)) Then
     vSQL = vSQL & " <> " & ActiveCell.value
   Else
     vSQL = vSQL & " not like '%" & ActiveCell.value & "%'"
   End If
   
   Call p_MoveActiveCellToHeader
   Call p_CreateTextBox("SqlQ", "" & vSQL)
   Call p_SqlQuery(ActiveSheet)
 Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " p_SQLRemoveOnly ")
   
End Sub
 
Sub p_StoreSQLHistory()
On Error GoTo ErrorHandler

Dim vSQLh
  vSQLh = getTextBoxValue("SqlHST")
  
  Call p_CreateTextBox("SqlHST", vSQLh & "#%#" & getTextBoxValue("SqlQ"))
  
 Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " p_StoreSQLHistory ")
 
 
End Sub

Sub p_SQLundo(ByVal vIRibbonControl As IRibbonControl)
On Error GoTo ErrorHandler
 Dim vSQL, i
 Dim vSqlArr() As String
 
   vSQL = getTextBoxValue("SqlHST")
   vSqlArr = Split(vSQL, "#%#")
   If UBound(vSqlArr) > 2 Then
        vSQL = vSqlArr(UBound(vSqlArr) - 1)
        
         Call p_CreateTextBox("SqlQ", "" & vSQL)
         
        vSQL = ""
        
        For i = 1 To (UBound(vSqlArr) - 2)
          vSQL = vSQL & "#%#" & vSqlArr(i)
        Next
        
        Call p_CreateTextBox("SqlHST", "" & vSQL)
        Call p_MoveActiveCellToHeader
        Call p_SqlQuery(ActiveSheet)
  End If
 Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " p_SQLundo ")
End Sub

Sub p_SqlQuery(vCurrSheet As Object)
On Error GoTo ErrorHandler
  
 If vCurrSheet Is Nothing Then
     MsgBox "Active sheet is not determinated ", vbExclamation
     Exit Sub
 End If
 
 Dim vSQL
 vSQL = getTextBoxValue("SqlQ")
  If (InStr(UCase(vSQL), "SELECT") > 0) Then
  
 Call p_setExcelCalcOff
  vCurrSheet.Cells(1, 1).Select
  
 Call p_StoreSQLHistory
   
 Call getDataFromUrl(getURLFromSheetWithID(vCurrSheet), "" & vSQL)
  
 
 Call p_setExcelCalcOn
  End If
 

 Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, " p_SqlQuery ")

End Sub

