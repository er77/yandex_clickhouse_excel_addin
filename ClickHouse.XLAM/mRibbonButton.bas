Attribute VB_Name = "mRibbonButton"
 Option Explicit
  
  Function f_SQLgetEnabled() As Boolean
     f_SQLgetEnabled = False
     Dim vSqlQ
      vSqlQ = getTextBoxValue("SQLConnectQ")
    If Len(vSqlQ) > 10 Then
        f_SQLgetEnabled = True
    End If
 End Function
 

Sub p_SQLgetEnabled(ByVal vIRibbonControl As IRibbonControl, ByRef vReturnValue)
     vReturnValue = f_SQLgetEnabled
     
     Select Case vIRibbonControl.ID
    Case "b_SQLAbout"
            vReturnValue = True
    Case "b_SQLCalculation"
             vReturnValue = False
    Case "b_SQLCellComments"
            vReturnValue = False
    Case "b_SQLConnections"
            vReturnValue = False
    Case "b_SQLCopySheet"
            vReturnValue = True
    Case "b_SQLDisconnect"
            vReturnValue = False
    Case "b_SQLHideTextBox"
            vReturnValue = False
    Case "b_SQLMemberInfo"
            vReturnValue = False
    Case "b_SQLMemberSelect"
           vReturnValue = False
    Case "b_SQLPivot"
            vReturnValue = False
    Case "b_SQLQueryDesigner"
            vReturnValue = False
    Case "b_SQLSheetInfo"
            vReturnValue = False
    Case "b_SQLQuickConnect"
            vReturnValue = True
    Case "b_SQLredo"
            vReturnValue = False
    Case "b_SQLsetAliasTable"
            vReturnValue = False
    Case "b_SQLShowTextBox"
            vReturnValue = False
    Case "b_SQLSubmitData"
            vReturnValue = False
    Case "b_SQLtest"
            vReturnValue = False
    Case "b_SQLZoomIn"
            vReturnValue = False
    Case "b_SQLZoomOut"
            vReturnValue = False
    Case "grp_SQLRData"
            vReturnValue = False
    Case "grp_SQLRefresh"
            vReturnValue = False
 Case "mn_Supr"
            vReturnValue = False
 Case "mn_Zoom"
            vReturnValue = False
 Case "mn_Selection"
            vReturnValue = False
 Case "mn_Show"
            vReturnValue = False
 Case "mn_Intend"
            vReturnValue = False
   End Select
      
 End Sub
 
 
   Public Sub getSQLlabel(ByVal vIRibbonControl As IRibbonControl, ByRef label)
 
    Dim vReturnValue
        vReturnValue = ""
      Select Case vIRibbonControl.ID
    Case "b_SQLAbout"
            vReturnValue = "About"
    Case "b_SQLCalculation"
            vReturnValue = "Calculus"
    Case "b_SQLCellComments"
            vReturnValue = "Cell Coments"
    Case "b_SQLConnections"
            vReturnValue = "b_SQLConnections"
    Case "b_SQLCopySheet"
            vReturnValue = "Copy Sheet"
    Case "b_ExcelCreatePivot"
            vReturnValue = "Excel Pivot"
    Case "b_SQLDisconnect"
            vReturnValue = "Disconnect"
    Case "b_SQLHideTextBox"
            vReturnValue = "Hide Text Box "
    Case "b_SQLKeepOnly"
            vReturnValue = "Keep Only"
    Case "b_SQLMemberInfo"
            vReturnValue = "Member Info"
    Case "b_SQLMemberSelect"
            vReturnValue = "Member Select"
    Case "b_SQLExecStored"
            vReturnValue = "SQL Editor"
    Case "b_SQLPivot"
            vReturnValue = "Pivot"
    Case "b_SQLQueryDesigner"
            vReturnValue = "Designer"
    Case "b_SQLQuickConnect"
            vReturnValue = "Quick Connect"
    Case "b_SQLredo"
            vReturnValue = "Redo"
    Case "b_SQLRemoveOnly"
            vReturnValue = "Remove Only"
    Case "b_SQLRetrieve"
            vReturnValue = "Retrieve Data"
    Case "b_RefreshAll"
            vReturnValue = "Retrieve All"
    Case "b_SQLsetAliasTable"
            vReturnValue = "Alias"
    Case "b_SQLSheetInfo"
            vReturnValue = "Options"
    Case "b_SQLShowTextBox"
            vReturnValue = "Show Text Box"
    Case "b_SQLSubmitData"
            vReturnValue = "Submit Data"
    Case "b_SQLtest"
            vReturnValue = "Test Dev"
    Case "b_SQLundo"
            vReturnValue = "Undo"
    Case "b_SQLZoomIn"
            vReturnValue = "Zoom In"
    Case "b_SQLZoomOut"
            vReturnValue = "Zoom Out"
    Case "grp_SQLRData"
            vReturnValue = "Additionals"
    Case "grp_SQLRefresh"
            vReturnValue = "SQL Tools"
    Case "mnu_SQLAddOption"
            vReturnValue = "More"


           
    End Select
     
    label = vReturnValue
  End Sub


 
