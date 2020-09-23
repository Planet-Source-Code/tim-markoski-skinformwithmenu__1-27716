VERSION 5.00
Begin VB.Form frmMainCode 
   BorderStyle     =   0  'None
   Caption         =   "Skin Form with Menu"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skin Form with Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   2025
   End
   Begin VB.Label lblFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&File"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   600
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   840
      ToolTipText     =   "Maximize (Disabled)"
      Top             =   1560
      Width           =   195
   End
   Begin VB.Image imgTitleRestore 
      Height          =   195
      Left            =   3000
      Picture         =   "frmMainCode.frx":0000
      ToolTipText     =   "Maximize (Disabled)"
      Top             =   2520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleMaximize 
      Height          =   195
      Left            =   3000
      Picture         =   "frmMainCode.frx":024A
      ToolTipText     =   "Maximize (Disabled)"
      Top             =   2280
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   3000
      Picture         =   "frmMainCode.frx":0494
      ToolTipText     =   "Minimize"
      Top             =   2040
      Width           =   195
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   3000
      Picture         =   "frmMainCode.frx":06DE
      ToolTipText     =   "Close"
      Top             =   1800
      Width           =   195
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   3000
      Picture         =   "frmMainCode.frx":0928
      ToolTipText     =   "Help"
      Top             =   1560
      Width           =   195
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   2280
      Picture         =   "frmMainCode.frx":0B72
      Top             =   1560
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   1920
      Picture         =   "frmMainCode.frx":12BC
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   2280
      Picture         =   "frmMainCode.frx":1A06
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   2640
      Picture         =   "frmMainCode.frx":2150
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   1560
      Picture         =   "frmMainCode.frx":289A
      Top             =   1560
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   1920
      Picture         =   "frmMainCode.frx":2FE4
      Top             =   1560
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   1560
      Picture         =   "frmMainCode.frx":372E
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   2640
      Picture         =   "frmMainCode.frx":3E78
      Top             =   1560
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skin Form with Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   4170
   End
End
Attribute VB_Name = "frmMainCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_KeyPress(p_intKeyAscii As Integer)
    ' Comments  :
    ' Parameters: p_intKeyAscii -
    ' Modified  :
    '
    ' --------------------------------------------------
    
    On Error GoTo PROC_ERR
    
    
    Select Case p_intKeyAscii
            
        Case 70, 102
            lblFile_Click
            
            
        Case Else
            
            ' Do nothing
            
    End Select
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
    
End Sub


Private Sub Form_Load()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    

    MakeWindow Me, False
    
    Load frmMenuForm
    
    ' Make the Maximize/Restore button have the Maximize image
    imgTitleMaxRestore.Picture = imgTitleMaximize.Picture
    

    LoadSkinz Me
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
    
End Sub


Private Sub Form_Unload(p_intCancel As Integer)
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
    Set frmMainCode = Nothing
    Unload frmMenuForm
    End
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
End Sub

Private Sub imgTitleClose_Click()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
Unload Me
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Private Sub imgTitleHelp_Click()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
    'SendKeys "{F1}", True

    Mbox "Put code here for calling your help file.", vbInformation
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
End Sub

Private Sub imgTitleLeft_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
    ' Comments  :
    ' Parameters: p_intButton
    '             p_intShift
    '             p_sngX
    '             p_sngY -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
    DoDrag Me
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
    
End Sub


Private Sub imgTitleMain_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
    ' Comments  :
    ' Parameters: p_intButton
    '             p_intShift
    '             p_sngX
    '             p_sngY -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
    DoDrag Me
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Private Sub imgTitleMinimize_Click()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
    Me.WindowState = vbMinimized
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Private Sub imgTitleRight_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
    ' Comments  :
    ' Parameters: p_intButton
    '             p_intShift
    '             p_sngX
    '             p_sngY -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
    DoDrag Me
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
    
End Sub


Private Sub lblFile_Click()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    PopupMenu frmMenuForm.mnuFile, , (lblFile.Left), (lblFile.Top + lblFile.Height)

    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Private Sub lblFile_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
    ' Comments  :
    ' Parameters: p_intButton
    '             p_intShift
    '             p_sngX
    '             p_sngY -
    ' Modified  :
    '
    ' --------------------------------------------------
    
    On Error GoTo PROC_ERR
    
    
    If p_intButton = 1 Then
        
        lblFile.BorderStyle = 1
    Else
        
        lblFile.BorderStyle = 0
        
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Private Sub lblFile_MouseUp(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
    ' Comments  :
    ' Parameters: p_intButton
    '             p_intShift
    '             p_sngX
    '             p_sngY -
    ' Modified  :
    '
    ' --------------------------------------------------
    
    On Error GoTo PROC_ERR
    
    
    If p_intButton = 1 Then
        
        lblFile.BorderStyle = 0
    Else
        
        lblFile.BorderStyle = 1
        
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
    
End Sub

