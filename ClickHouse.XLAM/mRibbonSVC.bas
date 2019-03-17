Attribute VB_Name = "mRibbonSVC"
 Option Explicit
  Private Const rbHandleProp = "ClickHouseRibbonHandleID"
  Public vIRibbonUI As IRibbonUI
  Public XLApp As CExcelEvents
  
  
#If VBA7 Then
    Public Declare PtrSafe Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#Else
    Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

#If VBA7 Then
Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
Function GetRibbon(ByVal lRibbonPointer As Long) As Object
#End If
        Dim objRibbon As Object
        CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
        Set GetRibbon = objRibbon
        Set objRibbon = Nothing
End Function

 Sub p_RefreshRibbonNow()
On Error Resume Next

   vIRibbonUI.Invalidate

    If Err.Number > 0 Then
           'vIRibbonUI.ActivateTab ("EssbaseAct")
            Set vIRibbonUI = GetRibbon(CLng(f_ReadGlobalProperty(rbHandleProp)))
            vIRibbonUI.Invalidate
         Err.Clear
         DoEvents
         Set XLApp = Nothing
         Set XLApp = New CExcelEvents
    End If
    
End Sub
Function getCountWB() As Integer
  
     getCountWB = ThisWorkbook.Sheets.Count
 End Function
 
Sub p_SQLOnRibbonLoad(vRibbon As IRibbonUI)
 On Error GoTo ErrorHandler
 Dim lngRibPtr
 Dim vCurrConnPrefix
 
 Dim cnt
 cnt = getCountWB()
 
 Application.MultiThreadedCalculation.Enabled = True
 Application.AutoRecover.Time = 7
 
     If Not (cnt = 0) Then
 
    Set vIRibbonUI = vRibbon

    lngRibPtr = ObjPtr(vRibbon)
    Call p_WriteGlobalProperty(rbHandleProp, lngRibPtr)
    
   End If
     Set XLApp = New CExcelEvents
l_exit:
    Exit Sub
ErrorHandler:
Call p_ErrorHandler(0, "p_OnRibbonLoad")
End Sub

 
 







