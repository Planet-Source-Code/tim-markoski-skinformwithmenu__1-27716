Attribute VB_Name = "MessageBox"
' Module     : MessageBox
' Description:
' Procedures : Mbox(p_strPrompt As String, Optional p_VbIcon As VbMsgBoxStyle, Optional p_strTitle As String)

' Modified   :
' 09/17/2001 TPM
'
' --------------------------------------------------
Option Explicit


Public Enum StandardIconEnum
    IDE_INFORMATION = 32516&       ' like vbInformation
    IDE_EXCLAMATION = 32515&    ' like vbExlamation
    IDE_CRITICAL = 32513&           ' like vbCritical
    IDE_QUESTION = 32514&       ' like vbQuestion
End Enum

Public Declare Function LoadStandardIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconNum As StandardIconEnum) As Long

Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long

Public Sub Mbox(p_strPrompt As String, Optional p_VbIcon As VbMsgBoxStyle, Optional p_strTitle As String)
    ' Comments  :
    ' Parameters: p_strPrompt
    '             p_VbIcon
    '             p_strTitle -
    ' Modified  :
    '
    ' --------------------------------------------------
    
    On Error GoTo PROC_ERR
    
    Dim oldPos                 As PointAPI
    Dim newPos                 As PointAPI
    Dim pIcon                  As StandardIconEnum
    Dim lngHIcon               As Long
    Dim lngOldWidth            As Long
    Dim lngOldHeight           As Long
    
    
    Select Case p_VbIcon
            
        Case vbInformation
            
            pIcon = IDE_INFORMATION
            
        Case vbExclamation
            
            pIcon = IDE_EXCLAMATION
            
        Case vbCritical
            
            pIcon = IDE_CRITICAL
            
        Case Else
            
            pIcon = IDE_EXCLAMATION
            
    End Select
    
    Load frmMsgBox
    
    lngOldWidth = (frmMsgBox.Width / Screen.TwipsPerPixelX)
    lngOldHeight = (frmMsgBox.Height / Screen.TwipsPerPixelY)
    oldPos.X = frmMsgBox.Left + lngOldWidth
    oldPos.Y = frmMsgBox.Top + lngOldHeight
    
    
    lngHIcon = LoadStandardIcon(0&, pIcon)
    
    Call DrawIcon(frmMsgBox.hDC, 24&, 40&, lngHIcon)
    
    If p_strTitle = vbNullString Then
        p_strTitle = App.Title
    End If
    
    frmMsgBox.lblMessage.Caption = p_strPrompt
    frmMsgBox.Width = (frmMsgBox.lblMessage.Left + frmMsgBox.lblMessage.Width) * Screen.TwipsPerPixelX
    frmMsgBox.Height = (frmMsgBox.lblMessage.Top + frmMsgBox.lblMessage.Height + frmMsgBox.cmdOK.Height + (40)) * Screen.TwipsPerPixelY
    newPos.X = frmMsgBox.lblMessage.Left + (frmMsgBox.Width / Screen.TwipsPerPixelX) - 0.75 * (Len(p_strPrompt))
    newPos.Y = frmMsgBox.Top + (frmMsgBox.Height)
    
    ResizeForm frmMsgBox, oldPos, newPos, 0
    
    frmMsgBox.cmdOK.Top = frmMsgBox.lblMessage.Top + frmMsgBox.lblMessage.Height + (frmMsgBox.cmdOK.Height)
    
    frmMsgBox.cmdOK.Left = ((frmMsgBox.Width / Screen.TwipsPerPixelX) - (frmMsgBox.cmdOK.Width)) / 2
    frmMsgBox.lblMessage.Caption = p_strPrompt
    frmMsgBox.lblTitle.Caption = p_strTitle
    frmMsgBox.Show 1
    frmMsgBox.Refresh
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
End Sub

