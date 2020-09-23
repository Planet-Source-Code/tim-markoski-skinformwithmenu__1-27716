VERSION 5.00
Begin VB.Form frmMenuForm 
   Caption         =   "Menu"
   ClientHeight    =   555
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(p_intCancel As Integer)
    ' Comments  :
    ' Parameters:  -
    ' Modified  :
    '
    ' --------------------------------------------------
    On Error GoTo PROC_ERR
    
    
    Set frmMenuForm = Nothing

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
    
End Sub

Private Sub mnuExit_Click()

Unload frmMainCode

End Sub

Private Sub mnuNew_Click()

Mbox "Put Code Here to Create a New File", vbInformation

End Sub

Private Sub mnuOpen_Click()

Mbox "Put Code Here to Open a File", vbInformation


End Sub

Private Sub mnuSave_Click()

Mbox "Put Code Here to Save an Open File", vbInformation

End Sub

Private Sub mnuSaveAs_Click()

IBox "Put Code Here to do a File/SaveAs.", App.Title, "Temp.txt"

End Sub
