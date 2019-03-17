Attribute VB_Name = "mToolsTextBox"
Option Explicit

Sub p_CreateTextBox(vNameOfTextBox As String, vText As String)
Call p_deleteAllTextBox(vNameOfTextBox)
  ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 50).Name = vNameOfTextBox
  ActiveSheet.Shapes(vNameOfTextBox).TextFrame.Characters.Text = vText
Call hideTextBox(vNameOfTextBox)
End Sub

Sub p_deleteAllTextBox(vNameOfTextBox As String)  ' delete All

Dim oTextBox As TextBox
 
For Each oTextBox In ActiveSheet.TextBoxes

  If InStr(UCase(oTextBox.Name), UCase(vNameOfTextBox)) > 0 Then
    oTextBox.Delete
  End If
Next oTextBox
Set oTextBox = Nothing
DoEvents
End Sub

Sub p_deleteTextBox(vNameOfTextBox As String)  ' delete if MDXq is more than 1

Dim oTextBox As TextBox
Dim i

i = 0

For Each oTextBox In ActiveSheet.TextBoxes

  If InStr(UCase(oTextBox.Name), UCase(vNameOfTextBox)) > 0 Then
   i = i + 1
  End If
  
  If i > 1 Then
    oTextBox.Delete
  End If
  
Next oTextBox
Set oTextBox = Nothing
DoEvents
End Sub

Public Function f_clearString(vDirtyString) As String
 Dim vTempString
   vTempString = Trim(vDirtyString)
   'vTempString = Replace(vTempString, vbCr, " ")
   'vTempString = Replace(vTempString, vbLf, " ")
   'vTempString = Replace(vTempString, vbCrLf, " ")
   vTempString = Replace(vTempString, ":", " ")
   'vTempString = Replace(vTempString, ">", " ")
   'vTempString = Replace(vTempString, "<", " ")
   vTempString = Replace(vTempString, "?", " ")
   vTempString = Replace(vTempString, "`", " ")
   vTempString = Replace(vTempString, "|", " ")
   vTempString = Replace(vTempString, "/", " ")
   vTempString = Replace(vTempString, "\", " ")
   vTempString = Replace(vTempString, "  ", " ")
   f_clearString = vTempString
End Function
Public Function isTextBoxPresent(vNameOfTextBox As String) As Boolean
On Error Resume Next
    isTextBoxPresent = False

    isTextBoxPresent = (Len(Trim(ActiveSheet.Shapes(vNameOfTextBox).TextFrame.Characters.Text)) > 0)
    If Err.Number > 0 Then
         Err.Clear
    End If
End Function

Function getTextBoxValue(vNameOfTextBox As String) As String
On Error Resume Next
getTextBoxValue = ""
    If isTextBoxPresent(vNameOfTextBox) Then
       getTextBoxValue = ActiveSheet.Shapes(vNameOfTextBox).TextFrame.Characters.Text
    End If
 Call p_deleteTextBox(vNameOfTextBox) ' delete if vNameOfTextBox  is more than 1
 Call hideTextBox(vNameOfTextBox)
    If Err.Number > 0 Then
         Err.Clear
    End If
End Function


Sub hideTextBox(vNameOfTextBox As String)
Dim vIsMyBox
On Error Resume Next
vIsMyBox = False

   If InStr(UCase(vNameOfTextBox), UCase("ConnectQ")) > 0 Then
     vIsMyBox = True
    End If
     
     If InStr(UCase(vNameOfTextBox), UCase("CalcQ")) > 0 Then
       vIsMyBox = True
     End If
     
     If InStr(UCase(vNameOfTextBox), UCase("SqlQ")) > 0 Then
        vIsMyBox = True
     End If
     
     If InStr(UCase(vNameOfTextBox), UCase("SqlHST")) > 0 Then
        vIsMyBox = True
     End If
     
    If InStr(UCase(vNameOfTextBox), UCase("MDXQ")) > 0 Then
        vIsMyBox = True
     End If
     
    If InStr(UCase(vNameOfTextBox), UCase("SQLConnectQ")) > 0 Then
        vIsMyBox = True
     End If
     
    If vIsMyBox Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 1
        ActiveSheet.Shapes(vNameOfTextBox).Height = 1
        ActiveSheet.Shapes(vNameOfTextBox).Left = 5000
        ActiveSheet.Shapes(vNameOfTextBox).Top = 5000
     End If
     
    If Err.Number > 0 Then
         Err.Clear
    End If
End Sub

Sub p_HideTextBox(vIRibbonControl As IRibbonControl)
Dim oTextBox As TextBox
Dim i

i = 0

For Each oTextBox In ActiveSheet.TextBoxes
 
  hideTextBox (oTextBox.Name)
  
Next oTextBox
Set oTextBox = Nothing
End Sub

Sub p_ShowTextBox(ByVal vIRibbonControl As IRibbonControl)
On Error Resume Next
Dim oTextBox As TextBox
Dim i

If Not (isTextBoxPresent("ConnectQ")) Then
  Call p_CreateTextBox("ConnectQ", "")
End If

If Not (isTextBoxPresent("MDXQ")) Then
  Call p_CreateTextBox("MDXQ", "")
End If

If Not (isTextBoxPresent("CalcQ")) Then
  Call p_CreateTextBox("CalcQ", "")
End If

If Not (isTextBoxPresent("SqlHST")) Then
  Call p_CreateTextBox("CalcQ", "")
End If
 
If Not (isTextBoxPresent("SQLConnectQ")) Then
  Call p_CreateTextBox("SQLConnectQ", "")
End If

If Not (isTextBoxPresent("SqlQ")) Then
  Call p_CreateTextBox("SqlQ", "")
End If

i = 0

For Each oTextBox In ActiveSheet.TextBoxes
 Call showTextBox(oTextBox.Name, i)
Next oTextBox



Set oTextBox = Nothing
End Sub

Sub showTextBox(vNameOfTextBox As String, ByVal i As Integer)

 On Error Resume Next
    '    ActiveSheet.Shapes(vNameOfTextBox).Width = 150
    '    ActiveSheet.Shapes(vNameOfTextBox).Height = 150
    '    ActiveSheet.Shapes(vNameOfTextBox).Left = 400
    '    ActiveSheet.Shapes(vNameOfTextBox).Top = 400
 
 
     If InStr(UCase(vNameOfTextBox), UCase("ConnectQ")) > 0 Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 100
        ActiveSheet.Shapes(vNameOfTextBox).Height = 20
        ActiveSheet.Shapes(vNameOfTextBox).Left = 10
        ActiveSheet.Shapes(vNameOfTextBox).Top = 10
      End If
      
     
     If InStr(UCase(vNameOfTextBox), UCase("CalcQ")) > 0 Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 150
        ActiveSheet.Shapes(vNameOfTextBox).Height = 150
       ActiveSheet.Shapes(vNameOfTextBox).Left = 70
        ActiveSheet.Shapes(vNameOfTextBox).Top = 70
     End If
     
     If InStr(UCase(vNameOfTextBox), UCase("SqlHST")) > 0 Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 155
        ActiveSheet.Shapes(vNameOfTextBox).Height = 155
        ActiveSheet.Shapes(vNameOfTextBox).Left = 105
        ActiveSheet.Shapes(vNameOfTextBox).Top = 105
     End If
     
     If InStr(UCase(vNameOfTextBox), UCase("SqlQ")) > 0 Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 150
        ActiveSheet.Shapes(vNameOfTextBox).Height = 150
        ActiveSheet.Shapes(vNameOfTextBox).Left = 100
        ActiveSheet.Shapes(vNameOfTextBox).Top = 100
     End If
     
     
     
    If InStr(UCase(vNameOfTextBox), UCase("MDXq")) > 0 Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 150
        ActiveSheet.Shapes(vNameOfTextBox).Height = 150
        ActiveSheet.Shapes(vNameOfTextBox).Left = 250
        ActiveSheet.Shapes(vNameOfTextBox).Top = 150
   With ActiveSheet.Shapes(vNameOfTextBox)
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
        .Solid
    End With

     End If
     
     If InStr(UCase(vNameOfTextBox), UCase("SQLConnectQ")) > 0 Then
        ActiveSheet.Shapes(vNameOfTextBox).Width = 150
        ActiveSheet.Shapes(vNameOfTextBox).Height = 150
        ActiveSheet.Shapes(vNameOfTextBox).Left = 350
        ActiveSheet.Shapes(vNameOfTextBox).Top = 150
     End If
     
     If Err.Number > 0 Then
         Err.Clear
    End If
End Sub
