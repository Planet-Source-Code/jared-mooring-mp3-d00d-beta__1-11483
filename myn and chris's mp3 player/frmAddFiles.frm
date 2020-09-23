VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Private Sub CmdOk_Click()
File1.Path = Dir1.Path
If File1.ListCount <> 0 Then
    For tel = 1 To File1.ListCount
        File1.ListIndex = tel - 1



        If Len(Dir1.Path) > 3 Then
            frmMain.lstfiles.AddItem Dir1.Path & "\" & File1.FileName
        Else
           'Exit For
            'MsgBox "You can't add a drive, only folders", vbOKOnly, "Error"
           'Exit Sub
        frmMain.lstfiles.AddItem Dir1.Path & File1.FileName
        End If
    Next tel
            Unload Me
Else
    MsgBox "No files were found in specific folder", vbOKOnly, "Error"
    Unload Me
End If
End Sub
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
