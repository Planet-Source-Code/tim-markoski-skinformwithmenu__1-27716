Attribute VB_Name = "InBox"
' Module     : InBox
' Description:
' Procedures : IBox(p_strPrompt As String, Optional p_strTitle As String, Optional p_strDefault As String)

' Modified   :
' 09/18/2001 TPM
'
' --------------------------------------------------
Option Explicit
Public g_strInpValue      As String

Public Function IBox(p_strPrompt As String, Optional p_strTitle As String, Optional p_strDefault As String) As String
    ' Comments  :
    ' Parameters: p_strPrompt
    '             p_strTitle
    '             p_strDefault -
    ' Returns   : String -
    ' Modified  :
    '
    ' --------------------------------------------------
    
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    
    On Error GoTo PROC_ERR
    
    
    Load frmInputBox
    
    frmInputBox.lblPrompt = p_strPrompt
    frmInputBox.lblTitle = p_strTitle
    frmInputBox.txtInput = p_strDefault
    
    frmInputBox.txtInput.SelStart = 0
    frmInputBox.txtInput.SelLength = 999
    
    frmInputBox.Show 1
    frmInputBox.Refresh
    
    IBox = g_strInpValue
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    Mbox Err.Description
    Resume PROC_EXIT
    
End Function

