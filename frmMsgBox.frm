VERSION 5.00
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Custom MsgBox with Icon"
   ClientHeight    =   2970
   ClientLeft      =   -15
   ClientTop       =   -75
   ClientWidth     =   5115
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMsgBox.frx":0000
   ScaleHeight     =   198
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   324
      Left            =   2040
      TabIndex        =   0
      Top             =   2400
      Width           =   1170
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1080
      MouseIcon       =   "frmMsgBox.frx":030A
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   600
      Width           =   45
   End
   Begin VB.Image imgTitleRestore 
      Height          =   195
      Left            =   6720
      Picture         =   "frmMsgBox.frx":0614
      Top             =   4200
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleMaximize 
      Height          =   195
      Left            =   6720
      Picture         =   "frmMsgBox.frx":085E
      Top             =   3960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   1080
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   6720
      Picture         =   "frmMsgBox.frx":0C1D
      Top             =   3240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   6720
      Picture         =   "frmMsgBox.frx":0E67
      Top             =   3480
      Width           =   195
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   6720
      Picture         =   "frmMsgBox.frx":10B1
      Top             =   3720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Machinist ToolBox Y2001®"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   3360
      Width           =   2505
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   6360
      Picture         =   "frmMsgBox.frx":12FB
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   6000
      Picture         =   "frmMsgBox.frx":1A45
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   5640
      Picture         =   "frmMsgBox.frx":218F
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   5280
      Picture         =   "frmMsgBox.frx":28D9
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   6360
      Picture         =   "frmMsgBox.frx":3023
      Top             =   3240
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   6000
      Picture         =   "frmMsgBox.frx":376D
      Top             =   3240
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   5640
      Picture         =   "frmMsgBox.frx":3EB7
      Top             =   3240
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   5280
      Picture         =   "frmMsgBox.frx":4601
      Top             =   3240
      Width           =   285
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module     : frmMsgBox
' Description:
' Procedures : cmdOK_Click()
'              Form_Load()
'              Form_Unload(p_intCancel As Integer)
'              imgTitleClose_Click()
'              imgTitleLeft_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
'              imgTitleMain_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
'              imgTitleRight_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)

' Modified   :
' 09/17/2001 TPM
'
' --------------------------------------------------
Option Explicit

Private Sub cmdOK_Click()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
    Unload frmMsgBox
    
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
    
    MakeWindow Me, True
    'AlwaysOnTop Me, True
    
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
    ' Parameters: p_intCancel -
    ' Modified  :
    '
    ' --------------------------------------------------
    
    On Error GoTo PROC_ERR
    
    
    Set frmMsgBox = Nothing
    
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

Private Sub lblTitle_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
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

