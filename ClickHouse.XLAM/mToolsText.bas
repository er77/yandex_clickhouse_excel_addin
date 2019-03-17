Attribute VB_Name = "mToolsText"
Option Explicit


Function getClearString(ByVal vCurrStr As String)
  Dim i As Long, vCodesToClean As Variant
  vCodesToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                       21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 96, 126, 127, 127, 129, 141, 143, 144, 157, 160)
  For i = LBound(vCodesToClean) To UBound(vCodesToClean)
    If InStr(vCurrStr, Chr(vCodesToClean(i))) Then vCurrStr = Replace(vCurrStr, Chr(vCodesToClean(i)), "")
  Next
  
  For i = 128 To 255
    If InStr(vCurrStr, Chr(i)) Then vCurrStr = Replace(vCurrStr, Chr(i), "")
  Next
    vCurrStr = Application.WorksheetFunction.Clean(vCurrStr)
    getClearString = Trim(vCurrStr)
End Function

Function getClearName(ByVal vCurrStr As String)
        Dim strPattern As String: strPattern = "[^a-zA-Z0-9]" 'The regex pattern to find special characters
        Dim strReplace As String: strReplace = "" 'The replacement for the special characters
        Dim regEx
        Set regEx = CreateObject("vbscript.regexp") 'Initialize the regex object
        ' Configure the regex object
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        ' Perform the regex replacement
        getClearName = regEx.Replace(vCurrStr, strReplace)
End Function



 

