VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlayList 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdRemoveFile 
      BackColor       =   &H000000FF&
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   -600
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   4800
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   600
      Top             =   5280
   End
   Begin VB.CommandButton cmdRemoveAll 
      BackColor       =   &H000000FF&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H000000FF&
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H000000FF&
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdListSave 
      BackColor       =   &H000000FF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdListOpen 
      BackColor       =   &H000000FF&
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "/\"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "\/"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton cmdAddDir 
      BackColor       =   &H000000FF&
      Caption         =   "Dir"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAddFile 
      BackColor       =   &H000000FF&
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H000000FF&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   855
   End
   Begin VB.Timer tmrSave 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   360
   End
   Begin VB.Timer tmrOpen 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4680
      Top             =   720
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   3720
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   2880
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstNames 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   3660
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   4215
   End
   Begin VB.ListBox lstfiles 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   3420
      Left            =   360
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblInfoTitle 
      Caption         =   "Label5"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label lblInfoArtist 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1920
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblSave 
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblOpen 
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmPlayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
cmdAddFile.Visible = True
cmdAddDir.Visible = True
End Sub

Private Sub cmdAddDir_Click()
frmAddDir.Show
End Sub

Private Sub cmdAddFile_Click()
'when clicked this shows the form addfile, so that the user can add their mp3's
frmAddFile.Show
End Sub

Private Sub cmdList_Click()
cmdListOpen.Visible = True
cmdListSave.Visible = True
End Sub

Private Sub cmdListOpen_Click()
On Error Resume Next
strFilter = "Playlist (*.mpd) |*.mpd"
cdOpen.Filter = strFilter
'open the common dialog in save mode
cdOpen.ShowOpen
'make sure that the retrieved filename is not a blank string
'
    If cdOpen.FileName <> "" Then
        'if it is not blank, open the file
        strFileName = cdOpen.FileName
        lblOpen.Caption = cdOpen.FileTitle
        OpenPathList cdOpen.FileName, frmPlayList
   
    End If
    
tmrOpen.Enabled = True
End Sub

Private Sub cmdListSave_Click()
    strFilter = "Playlist (*.mpd) |*.mpd"
    cdSave.Filter = strFilter
    
    'open the common dialog in save mode
    cdSave.ShowSave
    
    'make sure that the retrieved filename is not a blank string
    If cdSave.FileName <> "" Then
        'if it is not blank, open the file
        strFileName = cdSave.FileName
        lblSave.Caption = cdSave.FileTitle
        MousePointer = vbHourglass
        SavePathList cdSave.FileName, frmPlayList
        MousePointer = vbDefault
    End If
tmrSave.Enabled = True
End Sub

Private Sub cmdMoveDown_Click()
On Error Resume Next
Dim nItem As Integer
Dim nitem2 As Integer

With lstfiles
If lstfiles.ListIndex < 0 Then Exit Sub

nItem = lstfiles.ListIndex
nitem2 = lstNames.ListIndex
nitem3 = frmMain.lstfiles.ListIndex
nitem4 = frmMain.lstNames.ListIndex
If nItem = lstfiles.ListCount - 1 Then Exit Sub

lstfiles.AddItem lstfiles.Text, nItem + 2
lstNames.AddItem lstNames.Text, nitem2 + 2
frmMain.lstfiles.AddItem frmMain.lstfiles.Text, nitem3 + 2
frmMain.lstNames.AddItem frmMain.lstNames.Text, nitem4 + 2
lstfiles.RemoveItem nItem
lstNames.RemoveItem nitem2
frmMain.lstfiles.RemoveItem nitem3
frmMain.lstNames.RemoveItem nitem4

lstfiles.Selected(nItem + 1) = True
lstNames.Selected(nitem2 + 1) = True
frmMain.lstfiles.Selected(nitem3 + 1) = True
frmMain.lstNames.Selected(nitem4 + 1) = True

End With
End Sub

Private Sub cmdMoveUp_Click()
On Error Resume Next
Dim nItem As Integer
Dim nitem2 As Integer
Dim nitem3 As Integer
Dim nitem4 As Integer

With lstfiles
    If lstfiles.ListIndex < 0 Then Exit Sub
        nItem = lstfiles.ListIndex
        nitem2 = lstNames.ListIndex
        nitem3 = frmMain.lstfiles.ListIndex
        nitem4 = frmMain.lstNames.ListIndex
        If nItem = 0 Then Exit Sub
             If nitem2 = 0 Then Exit Sub

                lstfiles.AddItem lstfiles.Text, nItem - 1
                lstNames.AddItem lstNames.Text, nitem2 - 1
                frmMain.lstfiles.AddItem frmMain.lstfiles.Text, nitem3 - 1
                frmMain.lstNames.AddItem frmMain.lstNames.Text, nitem4 - 1
        
                lstfiles.RemoveItem nItem + 1
                lstNames.RemoveItem nitem2 + 1
                frmMain.lstfiles.RemoveItem nitem3 + 1
                frmMain.lstNames.RemoveItem nitem4 + 1
                
                lstfiles.Selected(nItem - 1) = True
                lstNames.Selected(nitem2 - 1) = True
                lstfiles.Selected(nItem - 1) = True
                lstNames.Selected(nitem2 - 1) = True
End With
End Sub

Private Sub cmdRemove_Click()
cmdRemoveFile.Visible = True
cmdRemoveAll.Visible = True
cmdAddFile.Visible = False
cmdAddDir.Visible = False
End Sub

Private Sub cmdRemoveAll_Click()
lstfiles.Clear
lstNames.Clear
frmMain.lstfiles.Clear
frmMain.lstNames.Clear
End Sub

Private Sub cmdRemoveFile_Click()
If lstfiles.ListIndex = -1 Then
    MsgBox "Please Select a file to remove from the Paylist", vbInformation, "mp3 d00d"
Else
    lstfiles.RemoveItem lstfiles.ListIndex
    lstNames.RemoveItem lstNames.ListIndex
    frmMain.lstfiles.RemoveItem frmMain.lstfiles.ListIndex
    frmMain.lstNames.RemoveItem frmMain.lstNames.ListIndex
End If
End Sub

Private Sub Command1_Click()
frmMain.chkPlayList.Value = 0
Me.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdListSave.Visible = False
cmdListOpen.Visible = False

cmdRemoveAll.Visible = False
cmdRemoveFile.Visible = False

cmdAddFile.Visible = False
cmdAddDir.Visible = False

End Sub

Private Sub lstNames_Click()
    
If lstNames.Selected(lstNames.ListIndex) = True Then
lstfiles.Selected(lstNames.ListIndex) = True
frmMain.lstfiles.Selected(lstNames.ListIndex) = True
frmMain.lstNames.Selected(lstNames.ListIndex) = True
txtPath = lstfiles.Text
Else
lstNames.Selected(lstNames.ListIndex) = False
lstfiles.Selected(lstfiles.ListIndex) = False
frmMain.lstNames.Selected(lstNames.ListIndex) = False
frmMain.lstfiles.Selected(lstfiles.ListIndex) = False
End If

End Sub

Private Sub lstNames_DblClick()
On Error GoTo bug
frmMain.MP.FileName = frmMain.txtPath.Text
'play the selected mp3

frmMain.MP.Play
'set the sliders(progress bar) maxamum number (the furthest it will go) to the duration
' of the song
frmMain.SldProgress.Max = frmMain.MP.Duration
'enable the timer so that it can start to count the time used in the song
frmMain.tmrTime.Enabled = True

'allow the pause label to become enabled
frmMain.cmdPause.Enabled = True

frmMain.tmrScroll.Enabled = True
'End If
frmMain.lbltitle.Caption = lstNames.Text
bug:
End Sub

Private Sub lstNames_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then 'keyboard event for the delete key
    Call cmdRemoveFile_Click
End If
    
End Sub

Private Sub tmrOpen_Timer()
OpenFileList "C:\windows\" & lblOpen.Caption, frmPlayList
tmrOpen.Enabled = False

End Sub

Private Sub tmrSave_Timer()
SaveFileList "C:\windows\" & lblSave.Caption, frmPlayList
tmrSave.Enabled = False
End Sub

