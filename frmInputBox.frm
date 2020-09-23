VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   0  'None
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   119
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1110
      Width           =   3750
   End
   Begin VB.CommandButton cmdEnter 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   324
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1170
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "&OK"
      Height          =   324
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   1170
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   4080
      Picture         =   "frmInputBox.frx":0000
      ToolTipText     =   "Help"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   4080
      Picture         =   "frmInputBox.frx":024A
      ToolTipText     =   "Close"
      Top             =   3840
      Width           =   195
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   4080
      Picture         =   "frmInputBox.frx":0494
      ToolTipText     =   "Minimize"
      Top             =   4080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   1920
      ToolTipText     =   "Maximize (Disabled)"
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleRestore 
      Height          =   195
      Left            =   4080
      Picture         =   "frmInputBox.frx":06DE
      ToolTipText     =   "Maximize (Disabled)"
      Top             =   4560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgTitleMaximize 
      Height          =   195
      Left            =   4080
      Picture         =   "frmInputBox.frx":0928
      ToolTipText     =   "Maximize (Disabled)"
      Top             =   4320
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   1800
      TabIndex        =   1
      Top             =   2640
      Width           =   60
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   3750
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   3360
      Picture         =   "frmInputBox.frx":0B72
      Top             =   3600
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   3720
      Picture         =   "frmInputBox.frx":12BC
      Top             =   3600
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   3000
      Picture         =   "frmInputBox.frx":1A06
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   3360
      Picture         =   "frmInputBox.frx":2150
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   3720
      Picture         =   "frmInputBox.frx":289A
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   285
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   2640
      Picture         =   "frmInputBox.frx":2FE4
      Top             =   3600
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   3000
      Picture         =   "frmInputBox.frx":372E
      Top             =   3600
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   2640
      Picture         =   "frmInputBox.frx":3E78
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   285
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module     : frmInputBox
' Description:
' Procedures : cmdEnter_Click(p_intIndex As Integer)
'              Form_Load()
'              Form_Unload(p_intCancel As Integer)
'              imgTitleLeft_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
'              imgTitleMain_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
'              imgTitleRight_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)
'              lblTitle_MouseDown(p_intButton As Integer, p_intShift As Integer, p_sngX As Single, p_sngY As Single)

' Modified   :
' 09/18/2001 TPM
'
' --------------------------------------------------
Option Explicit

Private Sub cmdEnter_Click(p_intIndex As Integer)
    ' Comments  :
    ' Parameters: p_intIndex -
    ' Modified  :
    '
    ' --------------------------------------------------
    
    On Error GoTo PROC_ERR
    
    
    Dim strValue As String
    
    
    Select Case p_intIndex
            
            
        Case 0
            
            strValue = txtInput
            
        Case 1
            
            strValue = vbNullString
            
    End Select
    
    g_strInpValue = strValue
    
    Unload Me
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Mbox Err.Description
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
    'AlwaysOnTop Me, True
    
    ' Make the Maximize/Restore button have the Maximize image
    '    imgTitleMaxRestore.Picture = imgTitleMaximize.Picture
    
    
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2 ' Center form vertically.
    
    txtInput.SelStart = 0
    txtInput.SelLength = 999
    
    LoadSkinz Me
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Mbox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Private Sub Form_Unload(p_intCancel As Integer)
    ' Comments  :
    ' Parameters: p_intCancel -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
    Set frmInputBox = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Mbox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Private Sub imgTitleClose_Click()
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
cmdEnter_Click 1
    
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
    Mbox Err.Description
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
    Mbox Err.Description
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
    Mbox Err.Description
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
    Mbox Err.Description
    Resume PROC_EXIT
    
    
End Sub

Private Sub txtInput_KeyPress(p_intKeyAscii As Integer)
    ' Comments  :
    ' Parameters: KeyAscii
    ' Modified  :
    '
    ' --------------------------------------------------
    
    On Error GoTo PROC_ERR
    
    If p_intKeyAscii = 13 Then
        
        Call cmdEnter_Click(0)
        
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Mbox Err.Description
    Resume PROC_EXIT
    
End Sub

