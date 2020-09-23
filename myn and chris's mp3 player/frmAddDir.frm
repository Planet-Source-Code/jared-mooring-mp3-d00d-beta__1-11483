VERSION 5.00
Begin VB.Form frmAddDir 
   BackColor       =   &H00000000&
   Caption         =   "Add Dir"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFF00&
      Caption         =   "Close"
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "Add"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   0
      Pattern         =   "*.mp3"
      TabIndex        =   3
      Top             =   6720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmAddDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
File1.path = Dir1.path
If File1.ListCount <> 0 Then
    For tel = 1 To File1.ListCount
        File1.ListIndex = tel - 1



        If Len(Dir1.path) > 3 Then
            frmMain.lstfiles.AddItem Dir1.path & "\" & File1.FileName
            frmMain.lstNames.AddItem File1.FileName
                 frmPlayList.lstfiles.AddItem Dir1.path & "\" & File1.FileName
            frmPlayList.lstNames.AddItem File1.FileName
        Else
           'Exit For
            'MsgBox "You can't add a drive, only folders", vbOKOnly, "Error"
           'Exit Sub
        frmMain.lstfiles.AddItem Dir1.path & File1.FileName
        frmMain.lstNames.AddItem File1.FileName
   
        End If
    Next tel
    

    
Else
    MsgBox "No mp3's were found in this folder", vbOKOnly + vbCritical, "Error"

End If
End Sub


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub

Private Sub Form_Load()
frmMain.Enabled = False
Ontop frmAddDir
frmAddDir.Left = frmMain.Left + 5
Dim GetDir
Dir1.path = GetSetting("mp3 d00d", "Properties", "Dir Path", "C:\")

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
SaveSetting "mp3 d00d", "Properties", "Dir Path", Dir1.path
End Sub
