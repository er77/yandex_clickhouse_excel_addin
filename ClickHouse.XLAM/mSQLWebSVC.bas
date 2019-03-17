Attribute VB_Name = "mSQLWebSVC"
Option Explicit

Function getURLwithQuery(vCurrentUrl, vCurrSQL) As String
   getURLwithQuery = vCurrentUrl & "/?query=" & vCurrSQL
End Function


Function pingYACH(ByVal vCurrentUrl) As Boolean

Dim oWinHttpRequest As WinHttp.WinHttpRequest
Dim sResult As String

On Error Resume Next

Set oWinHttpRequest = New WinHttp.WinHttpRequest
With oWinHttpRequest
    .Open "GET", vCurrentUrl, True
    .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    .Send ""
    .WaitForResponse
    sResult = oWinHttpRequest.Status
End With

 Set oWinHttpRequest = Nothing
 
pingYACH = False

 If Err.Number = 0 And sResult = "200" Then
    pingYACH = True
 End If

End Function


Function getYACHStatus(ByVal vCurrentUrl) As Boolean

Dim oWinHttpRequest As WinHttp.WinHttpRequest
Dim sResult As String

On Error Resume Next

Set oWinHttpRequest = New WinHttp.WinHttpRequest
With oWinHttpRequest
    .Open "GET", getURLwithQuery(vCurrentUrl, "SELECT+1"), True
    .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    .Send ""
    .WaitForResponse
    sResult = oWinHttpRequest.Status
End With

 Set oWinHttpRequest = Nothing
 
getYACHStatus = False

 If Err.Number = 0 And sResult = "200" Then
    getYACHStatus = True
 End If

End Function

Sub p_DeleteAllNamedRanges(ByVal xSTr As String)
'Update 20140314
Dim xName As Name
For Each xName In Application.ActiveWorkbook.Names
 If InStr(xName.Name, xSTr) > 0 Then
    xName.Delete
 End If
Next
End Sub

Sub p_RenameNamedRanges(ByVal xSTr As String)
'Update 20140314
Dim xName As Name
For Each xName In Application.ActiveWorkbook.Names
 If InStr(xName.Name, xSTr) > 0 Then
    xName.Name = xSTr
 End If
Next
End Sub


Sub p_deleteConnections(ByVal xSTr As String)
On Error Resume Next
    Dim qtb  As Variant
    For Each qtb In ThisWorkbook.Connections
       qtb.Delete
    Next
    
    Dim oQT As QueryTable
    For Each oQT In ActiveSheet.QueryTables
     If oQT.Refreshing Then
      oQT.CancelRefresh
     End If
     
     ActiveSheet.Range(Replace(oQT.Name, "'", "_")).Clear
     ActiveSheet.Names(Replace(oQT.Name, "'", "_")).Delete
     Call p_DeleteAllNamedRanges(Replace(oQT.Name, "'", "_"))
     
     ActiveWorkbook.Names(Replace(oQT.Name, "'", "_")).Delete
     
     ActiveSheet.QueryTables(ActiveSheet.Range(oQT.Name)).Delete
      Err.Clear
      ' oQT.WorkbookConnection.Delete
      ' oQT.Delete
    Next
    
  Call p_DeleteAllNamedRanges(xSTr)
  ActiveWorkbook.Names(xSTr).Delete

 
  
End Sub

Sub p_setUpClipboard(vSTR As String)
On Error Resume Next
Dim DataObj As New msforms.DataObject
   
    DataObj.SetText vSTR
    DataObj.PutInClipboard
End Sub
   
    
Sub getDataFromUrl(ByVal vCurrURL As String, vSQL As String) 'vCurrConnectString = vCurrConnectString & "&query=" & vSQL
On Error GoTo ErrorHandler
Dim vArr() As String
Dim vRangeName
vArr = Split(vCurrURL, "##")
 
'Call ReReadConnections
vCurrURL = "URL;" & vArr(1) '"http://192.168.56.101:9123/?query=SELECT 1" 'vCurrURL
 
 
If Not (InStr(UCase(vSQL), "LIMIT") > 0) Then
  vSQL = vSQL & " LIMIT 100000"
End If

If Not (InStr(UCase(vSQL), "FORMAT") > 0) Then
  vSQL = vSQL & " Format TabSeparatedWithNames"
End If

vSQL = Replace(vSQL, ";", "")
vSQL = vSQL & " ;"
Call p_setUpClipboard(vSQL)

vCurrURL = vCurrURL & "&query=" & vSQL

vRangeName = getCurrentTBRangeName
 Call p_deleteConnections(vRangeName)
    

With ActiveSheet.QueryTables.Add(Connection:=vCurrURL, destination:=vActiveCell)   'ActiveSheet.Cells(1, 1))
        '.PostText = "username=" & fArrQuickConnections(vArr(0), 1) & "&password=" & fArrQuickConnections(vArr(0), 2)
         .Name = vRangeName
        '.FieldNames = True
        '.RowNumbers = False
        '.FillAdjacentFormulas = False
        '.PreserveFormatting = True
         .RefreshOnFileOpen = False
         .BackgroundQuery = True
         .RefreshStyle = xlInsertDeleteCells
         .SavePassword = True
         .SaveData = True
         .AdjustColumnWidth = True
        '.RefreshPeriod = 0
        '.WebSelectionType = xlSpecifiedTables
         .WebFormatting = xlWebFormattingNone
         .WebTables = "2"
         .WebPreFormattedTextToColumns = True
        '.WebConsecutiveDelimitersAsOne = True
        '.WebSingleBlockTextImport = False
        '.WebDisableDateRecognition = False
        '.WebDisableRedirections = False
         .Refresh BackgroundQuery:=False
        
    .PostText = ""
    .RefreshStyle = xlOverwriteCells
     
    .Refresh
End With

Call p_RenameNamedRanges(vRangeName)

On Error Resume Next
 vActiveCell.AddComment
 Err.Clear
On Error GoTo ErrorHandler
 vActiveCell.Comment.Visible = False
 vActiveCell.Comment.Text Text:="URL: " & vCurrURL & Chr(13) & "SQL: " & vSQL
 
Exit Sub

ErrorHandler:
  MsgBox vCurrURL & Err.Number & Err.Description, vbCritical
  End
  
End Sub

