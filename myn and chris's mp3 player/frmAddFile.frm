VERSION 5.00
Begin VB.Form frmAddFile 
   BackColor       =   &H00000000&
   Caption         =   "Add FIle"
   ClientHeight    =   5430
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   7920
      Width           =   7935
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   735
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   4590
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   4965
      Left            =   4320
      MultiSelect     =   2  'Extended
      Pattern         =   "*.mp3"
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmAddFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Dim i As Integer

If File1.ListIndex = -1 Then Exit Sub

 For i = File1.ListCount - 1 To 0 Step -1
        If File1.Selected(i) = True Then
            frmMain.lstNames.AddItem basIni.GetFileTitle(File1.List(i))
            frmPlayList.lstNames.AddItem basIni.GetFileTitle(File1.List(i))
        End If
    Next i
    
For i = File1.ListCount - 1 To 0 Step -1

If File1.Selected(i) = True Then

frmMain.lstfiles.AddItem Dir1.path + "\" + File1.List(i)
frmPlayList.lstfiles.AddItem Dir1.path + "\" + File1.List(i)

  
End If

Next i
MsgBox "Your songs have been added", vbInformation, "Playlist"
            
End Sub

Private Sub Command1_Click()
Unload Me
frmMain.Show
End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
cmdAdd_Click
End Sub

Private Sub Form_Load()
Ontop frmAddFile
frmMain.Enabled = False
frmAddFile.Left = frmMain.Left + 5
Dir1.path = GetSetting("mp3 d00d", "Properties", "File Path", "C:\")
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
SaveSetting "mp3 d00d", "Properties", "File Path", Dir1.path
End Sub
