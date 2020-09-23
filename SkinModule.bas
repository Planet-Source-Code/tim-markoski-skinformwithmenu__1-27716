Attribute VB_Name = "SkinModule"
' Module     : SkinModule
' Description:
' Procedures : AlwaysOnTop(p_TheForm As Form, p_blnToggle As Boolean)
'              ChangeState(p_TheForm As Form)
'              DoDrag(p_TheForm As Form)
'              DoTransparency(p_TheForm As Form)
'              LoadSkinz(p_FrmSkin As Form)
'              MakeWindow(p_TheForm As Form, p_blnIsResizable As Boolean)
'              ResizeForm(p_TheForm As Form, p_OldCursorPos As PointAPI, p_NewCursorPos As PointAPI, p_intResizeMode As Integer)
'              SetStateBtn(p_TheForm As Form, p_lngNewState As Long)

' Modified   :
' 09/19/2001 TPM
'
' --------------------------------------------------

Option Explicit

' xx Used to set the shape of the form
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
' xx Used to create the rounded rectangle region
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' xx Used to make the form draggable
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' xx Also used to make the form draggable
Public Declare Function ReleaseCapture Lib "user32" () As Long
' xx Used to make the window always on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' xx Used to get the cursor position
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
' xx Various bits and pieces used by the above functions
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type PointAPI
    X As Long
    Y As Long
End Type

Dim m_intResizable As Integer

Public Sub AlwaysOnTop(p_TheForm As Form, p_blnToggle As Boolean)
    ' Comments  :
    ' Parameters: p_TheForm
    '             p_blnToggle -
    ' Modified  :
    '
    ' --------------------------------------------------
    '  TheForm:  The form you want to make always on top or not
    '  Toggle:   (True/False) - True for always on top, False for normal
    
    On Error GoTo PROC_ERR
    
    
    If p_blnToggle = True Then
        SetWindowPos p_TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    Else
        SetWindowPos p_TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    End If
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Mbox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Public Sub ChangeState(p_TheForm As Form)
    ' Comments  :
    ' Parameters: p_TheForm -
    ' Modified  :
    '
    ' --------------------------------------------------
    '  TheForm:  The form you want to change state (maximized, normal)
    
    On Error GoTo PROC_ERR
    
    
    If p_TheForm.WindowState = vbNormal Then
        p_TheForm.WindowState = vbMaximized
        p_TheForm!imgTitleMaxRestore.Picture = p_TheForm!imgTitleRestore.Picture
        MakeWindow p_TheForm, False
    Else
        p_TheForm.WindowState = vbNormal
        p_TheForm!imgTitleMaxRestore.Picture = p_TheForm!imgTitleMaximize.Picture
        MakeWindow p_TheForm, IIf(m_intResizable = 1, True, False)
    End If
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Mbox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Public Sub DoDrag(p_TheForm As Form)
    ' Comments  :
    ' Parameters: p_TheForm -
    ' Modified  :
    '
    ' --------------------------------------------------
    '  TheForm:  The form you want to start dragging
    
    On Error GoTo PROC_ERR
    
    
    If p_TheForm.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessage p_TheForm.hwnd, &HA1, 2, 0&
    End If
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Mbox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Public Sub DoTransparency(p_TheForm As Form)
    ' Comments  :
    ' Parameters: p_TheForm -
    ' Modified  :
    '
    ' --------------------------------------------------
    '  TheForm:  The form you want to be rounded rectangle shape
    
    On Error GoTo PROC_ERR
    
    
    Dim alngTempRegions(6) As Long
    Dim lngFormWidthInPixels As Long
    Dim lngFormHeightInPixels As Long
    Dim varA
    
    '  Convert the form's height and width from twips to pixels
    lngFormWidthInPixels = p_TheForm.Width / Screen.TwipsPerPixelX
    lngFormHeightInPixels = p_TheForm.Height / Screen.TwipsPerPixelY
    
    '  Make a rounded rectangle shaped region with the dimensions of the form
    varA = CreateRoundRectRgn(0, 0, lngFormWidthInPixels, lngFormHeightInPixels, 24, 24)
    
    '  Set this region as the shape for "TheForm"
    varA = SetWindowRgn(p_TheForm.hwnd, varA, True)
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Mbox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Public Sub LoadSkinz(p_FrmSkin As Form)
    ' Comments  :
    ' Parameters: p_FrmSkin -
    ' Modified  :
    '
    ' --------------------------------------------------
    
    On Error Resume Next
    
    
    Dim strSkin         As String
    Dim ctl             As Control
    
    strSkin = GetSetting(App.Title, "Options", "Current Skin", "Blue")
    
    
    For Each ctl In p_FrmSkin.Controls
        
        If TypeOf ctl Is Image Then
            
            If InStr(1, ctl.Name, "img", vbTextCompare) > 0 Then
                'Debug.Print ctl.Name
                ctl.Picture = LoadPicture(App.Path & "\Skinz\" & strSkin & "\" & ctl.Name & ".gif")
                ctl.Refresh

            End If
            
        End If
        
    Next
    
    
    p_FrmSkin.imgTitleMaxRestore.Picture = p_FrmSkin.imgTitleMaximize.Picture
    DoEvents
    
End Sub

Public Sub MakeWindow(p_TheForm As Form, p_blnIsResizable As Boolean)
    ' Comments  :
    ' Parameters: p_TheForm
    '             p_blnIsResizable -
    ' Modified  :
    '
    ' --------------------------------------------------
    '  TheForm:           The form you want to make graphical
    '  IsResizable:       (True/False) - True for resizable at runtime
    
    '  Declare some variables
    On Error GoTo PROC_ERR
    
    
    Dim lngFormWidth As Long
    Dim lngFormHeight As Long
    Dim intTemp As Integer
    
    '  Set the Resizable variable
    m_intResizable = IIf(p_blnIsResizable = True, 1, 0)
    
    '  Store the form's width and height in pixels in a variable
    lngFormWidth = (p_TheForm.Width / Screen.TwipsPerPixelX)
    lngFormHeight = (p_TheForm.Height / Screen.TwipsPerPixelY)
    
    '  Set various parameters of the form
    p_TheForm.BackColor = RGB(192, 192, 192)
    p_TheForm.Caption = p_TheForm!lblTitle.Caption
    
    '  Set the position of the title label
    p_TheForm!lblTitle.Left = 16
    p_TheForm!lblTitle.Top = 7
    
    '  Make the form "rounded rectangle" shaped (call to the sub below)
    DoTransparency p_TheForm
    
    '' xx Move the image blocks into place and stretch them accordingly
    With p_TheForm!imgTitleLeft
        .Top = 0
        .Left = 0
    End With
    
    With p_TheForm!imgTitleRight
        .Top = 0
        .Left = lngFormWidth - 19
    End With
    
    With p_TheForm!imgTitleMain
        .Top = 0
        .Left = 19
        .Width = lngFormWidth - 19
    End With
    
    With p_TheForm!imgWindowLeft
        .Top = 30
        .Left = 0
        .Height = lngFormHeight - 60
    End With
    
    With p_TheForm!imgWindowBottomLeft
        .Top = lngFormHeight - 30
        .Left = 0
    End With
    
    With p_TheForm!imgWindowBottom
        .Top = lngFormHeight - 30
        .Left = 19
        .Width = lngFormWidth - 38
    End With
    
    With p_TheForm!imgWindowBottomRight
        .Top = lngFormHeight - 30
        .Left = lngFormWidth - 19
    End With
    
    With p_TheForm!imgWindowRight
        .Top = 30
        .Left = lngFormWidth - 19
        .Height = lngFormHeight - 38
    End With
    
    '  Position the title buttons (close, minimize, help)
    With p_TheForm!imgTitleClose
        .Top = 8
        .Left = lngFormWidth - 22
    End With
    
    With p_TheForm!imgTitleMaxRestore
        .Top = 8
        .Left = lngFormWidth - 39
    End With
    
    With p_TheForm!imgTitleMinimize
        .Top = 8
        .Left = lngFormWidth - 56
    End With
    
    With p_TheForm!imgTitleHelp
        .Top = 8
        .Left = lngFormWidth - 73
    End With
    
    '' xx Position the resizing invisible images
    '   If IsResizable = True Then
    '       For Temp = 0 To 7
    '           TheForm!Resizer(Temp).Visible = True
    '       Next Temp
    '
    '       With TheForm!Resizer(0)
    '           .Top = 30
    '           .Left = 0
    '           .Height = FormHeight - 60
    '       End With
    '
    '       With TheForm!Resizer(1)
    '           .Top = 30
    '           .Left = FormWidth - 5
    '           .Height = FormHeight - 60
    '       End With
    '
    '       With TheForm!Resizer(2)
    '           .Top = 0
    '           .Left = 19
    '           .Width = FormWidth - 39
    '       End With
    '
    '       With TheForm!Resizer(3)
    '           .Top = FormHeight - 5
    '           .Left = 19
    '           .Width = FormWidth - 39
    '       End With
    '
    '       With TheForm!Resizer(4)
    '           .Top = FormHeight - 11
    '           .Left = FormWidth - 11
    '       End With
    '
    '       With TheForm!Resizer(5)
    '           .Top = FormHeight - 11
    '           .Left = 0
    '       End With
    '
    '       With TheForm!Resizer(6)
    '           .Top = 0
    '           .Left = FormWidth - 11
    '       End With
    '
    '       With TheForm!Resizer(7)
    '           .Top = 0
    '           .Left = 0
    '       End With
    '   End If
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Mbox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Public Sub ResizeForm(p_TheForm As Form, p_OldCursorPos As PointAPI, p_NewCursorPos As PointAPI, p_intResizeMode As Integer)
    ' Comments  :
    ' Parameters: p_TheForm
    '             p_OldCursorPos
    '             p_NewCursorPos
    '             p_intResizeMode -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error Resume Next
    
    '  TheForm:      The form you want to resize
    '  OldCursorPos: The old cursor position (MouseDown)
    '  NewCursorPos: The new cursor position (MouseUp)
    '  ResizeMode:   0 - Left side
    '                1 - Right side
    '                2 - Top side
    '                3 - Bottom side
    '                4 - Bottom right corner
    '                5 - Bottom left corner
    '                6 - Top right corner
    '                7 - Top left corner
    
    '  Declare some variables
    Dim varDifferenceX
    Dim varDifferenceY
    
    '  Put the difference between the first cursor pos and the second into variables
    varDifferenceX = (p_NewCursorPos.X - p_OldCursorPos.X) * Screen.TwipsPerPixelX
    varDifferenceY = (p_NewCursorPos.Y - p_OldCursorPos.Y) * Screen.TwipsPerPixelY
    
    '  Determine which resizing mode (above) has been called and resize accordingly
    Select Case p_intResizeMode
        Case 0
            p_TheForm.Move p_TheForm.Left + varDifferenceX, p_TheForm.Top, p_TheForm.Width - varDifferenceX, p_TheForm.Height
        Case 1
            p_TheForm.Move p_TheForm.Left, p_TheForm.Top, p_TheForm.Width + varDifferenceX, p_TheForm.Height
        Case 2
            p_TheForm.Move p_TheForm.Left, p_TheForm.Top + varDifferenceY, p_TheForm.Width, p_TheForm.Height - varDifferenceY
        Case 3
            p_TheForm.Move p_TheForm.Left, p_TheForm.Top, p_TheForm.Width, p_TheForm.Height + varDifferenceY
        Case 4
            p_TheForm.Move p_TheForm.Left, p_TheForm.Top, p_TheForm.Width + varDifferenceX, p_TheForm.Height + varDifferenceY
        Case 5
            p_TheForm.Move p_TheForm.Left + varDifferenceX, p_TheForm.Top, p_TheForm.Width - varDifferenceX, p_TheForm.Height + varDifferenceY
        Case 6
            p_TheForm.Move p_TheForm.Left, p_TheForm.Top + varDifferenceY, p_TheForm.Width + varDifferenceX, p_TheForm.Height - varDifferenceY
        Case 7
            p_TheForm.Move p_TheForm.Left + varDifferenceX, p_TheForm.Top + varDifferenceY, p_TheForm.Width - varDifferenceX, p_TheForm.Height - varDifferenceY
    End Select
    
    '  Check to see if the form has been resized below the minimum size
    '    If TheForm.Width < 57 * Screen.TwipsPerPixelX Then TheForm.Width = 57 * Screen.TwipsPerPixelX
    '    If TheForm.Height < 90 * Screen.TwipsPerPixelY Then TheForm.Height = 90 * Screen.TwipsPerPixelY
    
    '  After resizing the form, make the form "rounded rectangle" shaped
    MakeWindow p_TheForm, True
End Sub

Public Sub SetStateBtn(p_TheForm As Form, p_lngNewState As Long)
    ' Comments  :
    ' Parameters: p_TheForm
    '             p_lngNewState -
    ' Modified  :
    '
    ' --------------------------------------------------
    '  TheForm:  The form you want to set the Max/Restore button
    '  NewState: A vbConstant for the state
    
    On Error GoTo PROC_ERR
    
    
    If p_lngNewState <> vbNormal Then
        p_TheForm!imgTitleMaxRestore.Picture = p_TheForm!imgTitleRestore.Picture
    Else
        p_TheForm!imgTitleMaxRestore.Picture = p_TheForm!imgTitleMaximize.Picture
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Mbox Err.Description
    Resume PROC_EXIT
    
    
End Sub

