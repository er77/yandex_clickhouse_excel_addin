Attribute VB_Name = "mRibbonOPT"
 Option Explicit
 
 Public vHSV_SUPPRESSCOLUMNS_MISSING As Boolean
 Public vHSV_SUPPRESSCOLUMNS_ZEROS As Boolean
 Public vHSV_SUPPRESS_MISSINGBLOCKS As Boolean
 Public vHSV_SUPPRESSROWS_MISSING As Boolean
 Public vHSV_SUPPRESSROWS_ZEROS As Boolean
 Public vHSV_MEMBER_DISPLAY  As Integer
 Public vHSV_ZOOMIN As Integer
 Public vHSV_ANCESTOR_POSITION As Integer

 Public vHSV_MISSING_LABEL As String
 Public vHSV_INDENTATION As Integer
 Public vHSV_INCLUDE_SELECTION As Boolean
 Public vIsDeleteOrphans As Boolean
 Public vIsSuppressOnPivot As Boolean
 
 Public vIsUseNameDefault  As Boolean
 
 
 Public vIsAlreadyHide As Boolean
 
 
 Public vModeAnalyse As Integer
 
 Public X As Long
 
 Sub p_setAliasOptions(vIRibbonControl As IRibbonControl)
             
    
 End Sub

 Sub p_onActionINT(vIRibbonControlID As String, ByVal vSelectedValue)
On Error GoTo ErrorHandler
Dim isEssOption

isEssOption = True
 

    Select Case vSelectedValue
     '   Case "mn_Supr0"
     '       vHSV_SUPPRESS_MISSINGBLOCKS = True
     '       vHSV_SUPPRESSROWS_MISSING = False

        Case "mn_Supr1"
            vHSV_SUPPRESSROWS_MISSING = True
            vHSV_SUPPRESS_MISSINGBLOCKS = True
            vHSV_SUPPRESSROWS_ZEROS = True
            vHSV_SUPPRESSCOLUMNS_MISSING = True
            vHSV_SUPPRESSCOLUMNS_ZEROS = True
            
        Case "mn_Supr2"
            vHSV_SUPPRESSROWS_MISSING = True
            vHSV_SUPPRESS_MISSINGBLOCKS = False
            vHSV_SUPPRESSROWS_ZEROS = False
             Call p_mnMissInit
             
        Case "mn_Supr6"
            vHSV_SUPPRESS_MISSINGBLOCKS = False
            vHSV_SUPPRESSROWS_MISSING = False
              Call p_mnMissInit
              
        Case "mn_SuprClmn2"
            vHSV_SUPPRESSCOLUMNS_MISSING = True

        Case "mn_SuprClmn6"
            vHSV_SUPPRESSCOLUMNS_MISSING = False

        Case "mn_Zoom0"
            vHSV_ZOOMIN = 2

        Case "mn_Zoom1"
            vHSV_ZOOMIN = 0

        Case "mn_Zoom2"
            vHSV_ZOOMIN = 1

        Case "mn_Intend0"
            vHSV_INDENTATION = 0

        Case "mn_Intend1"
            vHSV_INDENTATION = 1

        Case "mn_Intend2"
            vHSV_INDENTATION = 2
            
        Case "mn_Alias0"
             
            Call p_setExcelCalcOff
             vIsUseNameDefault = True
            
            Call p_setExcelCalcOn
             
        Case "mn_Alias1"
             
            Call p_setExcelCalcOff
             vIsUseNameDefault = False
             
            Call p_setExcelCalcOn
             
        Case "mn_Alias2"
            vIsUseNameDefault = False
             


        Case "mn_Show1"
            vHSV_MEMBER_DISPLAY = 2
                
        Case "mn_Show2"
            vHSV_MEMBER_DISPLAY = 1
            
      Case "mn_SubTot0"
            vHSV_ANCESTOR_POSITION = 0

        Case "mn_SubTot1"
            vHSV_ANCESTOR_POSITION = 1

         Case "mn_Selection0"
            vHSV_INCLUDE_SELECTION = True
 
        Case "mn_Selection1"
            vHSV_INCLUDE_SELECTION = False
 

         Case "mn_DelOrp0"
            vIsDeleteOrphans = True
            
        Case "mn_DelOrp1"
            vIsDeleteOrphans = False
 
 
            
         Case "mn_Mode0"
            vModeAnalyse = 0
            p_RefreshRibbonNow
            isEssOption = False
            
         Case "mn_Mode1"
            vModeAnalyse = 1
            p_RefreshRibbonNow
            isEssOption = False
            
         Case "mn_Mode2"
            vModeAnalyse = 2
            p_RefreshRibbonNow
            isEssOption = False

    End Select

  Call p_WriteGlobalProperty(vIRibbonControlID, vSelectedValue)

    If ActiveSheet Is Nothing Then
        GoTo l_exit
    End If
 
    If isEssOption Then
      Call p_setCurrentOptions(vIRibbonControlID, vSelectedValue)
    End If

l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(X, "p_onAction")
End Sub
 
  Sub pb_onAction_supr(vIRibbonControl As IRibbonControl, ByRef returnedVal)
 
 
   
 
End Sub
 
 Sub p_onAction_supr(vIRibbonControl As IRibbonControl, ByVal vSelectedValue, Optional vOptional)
 
 
 
End Sub
 
 Sub p_onAction(vIRibbonControl As IRibbonControl, ByVal vSelectedValue, Optional vOptional)
 
 

Call p_onActionINT(vIRibbonControl.ID, vSelectedValue)
 
End Sub
 Sub p_onChange(vIRibbonControl As IRibbonControl, ByRef vSelectedValue As Variant)
On Error GoTo ErrorHandler
    Select Case vIRibbonControl.ID
        Case "mn_MissLabel"
           vHSV_MISSING_LABEL = vSelectedValue
            
          
    End Select
l_exit:
    Exit Sub
ErrorHandler:
     Call p_ErrorHandler(X, "p_onChange")
End Sub
Sub p_setCurrentOptions(vIRibbonControlID As String, ByVal vSelectedValue)

   
     Select Case vIRibbonControlID
        Case "mn_Supr"
            
            
        Case "mn_Zoom"
             
            
        Case "mn_Intend"
             

        Case "mn_Show"
              
              
        Case "mn_SubTot"
              
        Case "mn_Selection"
            
    End Select
End Sub

Sub p_getSelectedItemIDINT(vIRibbonControlID As String, ByRef itemID As Variant)
Dim isEssOption

    If ActiveSheet Is Nothing Then

      If Not vIsAlreadyHide Then
       vIsAlreadyHide = True
       
       
      End If
      
       
       Exit Sub
    End If
    
   itemID = Null
 'DoEvents
 
   itemID = f_ReadGlobalProperty(vIRibbonControlID)
   
   isEssOption = True
   If (IsNull(itemID)) Then
     Select Case vIRibbonControlID
        Case "mn_Supr"
              
              itemID = "mn_Supr6"
              
               
              
               
               
        Case "mn_SuprClmn"
              itemID = "mn_SuprClmn6"
              vHSV_SUPPRESSCOLUMNS_MISSING = False
               
        Case "mn_Zoom"
             itemID = "mn_Zoom0"
             vHSV_ZOOMIN = 2
             
        Case "mn_Intend"
              itemID = "mn_Intend0"
                vHSV_INDENTATION = 0
                
        Case "mn_Show"
              itemID = "mn_Show1"
              vHSV_MEMBER_DISPLAY = 2
              
        Case "mn_AddSupr"
              itemID = "mn_Supr5"
              vIsSuppressOnPivot = False
              
        Case "mn_SubTot"
               itemID = "mn_SubTot0"
               vHSV_ANCESTOR_POSITION = 0
               
        Case "mn_Mode"
               itemID = "mn_Mode0"
               vModeAnalyse = 0
               
        ' Case "mn_Alias"
        '
        '    If vIsUseNameDefault Then
        '      itemID = "mn_Alias0"
        '    Else
        '      itemID = "mn_Alias1"
        '    End If
               
         Case "mn_Selection"
            itemID = "mn_Selection0"
            vHSV_INCLUDE_SELECTION = True
            
         Case "mn_DelOrp"
            itemID = "mn_DelOrp0"
            vIsDeleteOrphans = True
            
    End Select
    
    If isEssOption Then
      Call p_setCurrentOptions(vIRibbonControlID, itemID)
    End If
Else
     If InStr(vIRibbonControlID, "mn_Show") = 0 Then
        Call p_onActionINT(vIRibbonControlID, itemID)
     End If
 End If
    
End Sub

Sub p_getSelectedItemID(vIRibbonControl As IRibbonControl, ByRef itemID As Variant, Optional vOptional)
 

Call p_getSelectedItemIDINT(vIRibbonControl.ID, itemID)
    
  
     
End Sub

 
Sub p_mnMissInit()
 
    
End Sub
 
 Sub p_SetOptionSuppress()
    
 End Sub
             
Sub p_restoreOptions()
Dim itemID
 'ActiveSheet.Cells(1, 1).Select
  
End Sub
 

Sub p_SetOption(vStrOption As Variant, vOptionValue As Variant)
 
 
End Sub
Function getCountWB() As Integer
  
     getCountWB = ThisWorkbook.Sheets.Count
 End Function

Sub p_SetAllOptions()

      
           
l_exit:
     Exit Sub
ErrorHandler:
 Call p_ErrorHandler(X, "p_SetAllOptions")
     
End Sub

