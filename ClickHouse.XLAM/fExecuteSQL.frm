VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fExecuteSQL 
   Caption         =   "Change and Execute SQL"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   OleObjectBlob   =   "fExecuteSQL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fExecuteSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bExecuteSQL_Click()
    Unload Me
   Call p_CreateTextBox("SqlQ", "" & Me.mdxTextBox.Text)
   Call p_SqlQuery(ActiveSheet)

End Sub

Private Sub bSaveSQL_Click()
    Call p_CreateTextBox("SqlQ", "" & Me.mdxTextBox.Text)
    Unload Me
End Sub

Private Sub fCancel_Click()
 Unload Me
End Sub

Private Sub UserForm_Initialize()
 On Error GoTo ErrorHandler
  Me.Describes.Text = "Enter a valid SQL statement.Successfull execution of SQL statment will overwrite currently selected worksheet."
  Me.mdxTextBox.Text = getTextBoxValue("SqlQ")
   
l_exit:
    Exit Sub
ErrorHandler:
 Call p_ErrorHandler(0, "ExecuteSQL: UserForm_Initialize ")
  Unload Me
End Sub

 
