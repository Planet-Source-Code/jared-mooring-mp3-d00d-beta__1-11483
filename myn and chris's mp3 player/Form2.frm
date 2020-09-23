VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7605
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optShuffle 
      Caption         =   "S"
      Height          =   255
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Shuffle"
      Top             =   480
      Width           =   255
   End
   Begin VB.OptionButton optNormal 
      Caption         =   "N"
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Normal Play"
      Top             =   480
      Width           =   255
   End
   Begin VB.OptionButton optRepeat 
      Caption         =   "R"
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Repeat"
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   29
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   28
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   27
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   26
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   ";"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   25
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      Top             =   1320
      Width           =   375
   End
   Begin VB.Timer tmrOpen 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   3360
   End
   Begin VB.Timer tmrSave 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   3120
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   3600
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   2760
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkPlayList 
      Caption         =   "PL"
      Height          =   255
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "^"
      Height          =   255
      Left            =   3720
      TabIndex        =   20
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "_"
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "\/"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5040
      Width           =   375
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "/\"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4200
      Width           =   375
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Shuffle"
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "R"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   360
      Width           =   3495
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6360
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   3855
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   3  'Dot
         X1              =   1560
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lbltitle 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6720
      Top             =   1440
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1080
      Width           =   8175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   2040
   End
   Begin MSComctlLib.Slider SldProgress 
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   344
      _Version        =   393216
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   195
      Left            =   2760
      TabIndex        =   6
      ToolTipText     =   "Volume"
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   344
      _Version        =   393216
      Max             =   2500
      TickStyle       =   3
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
      Height          =   3420
      Left            =   480
      TabIndex        =   0
      Top             =   3240
      Width           =   5055
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
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label lblPlaylist 
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblOpen 
      Height          =   495
      Left            =   4440
      TabIndex        =   23
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblSave 
      Height          =   495
      Left            =   4200
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   255
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   8
      X2              =   392
      Y1              =   464
      Y2              =   464
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   8
      X2              =   8
      Y1              =   192
      Y2              =   464
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   392
      X2              =   392
      Y1              =   192
      Y2              =   464
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8
      X2              =   392
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Label lblInfoArtist 
      Caption         =   "Label1"
      Height          =   495
      Left            =   7680
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblInfoTitle 
      Caption         =   "Label5"
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin MediaPlayerCtl.MediaPlayer MP 
      Height          =   615
      Left            =   -120
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -40
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuNext 
         Caption         =   "Next"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPrevious 
         Caption         =   "Previous"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
         Shortcut        =   ^D
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile2 
         Caption         =   "File"
         Begin VB.Menu mnuAdd 
            Caption         =   "Add"
            Begin VB.Menu mnuFile3 
               Caption         =   "File"
               Shortcut        =   ^O
            End
            Begin VB.Menu mnuDir 
               Caption         =   "Dir"
               Shortcut        =   ^P
            End
         End
         Begin VB.Menu mnuRemoveAll 
            Caption         =   "Remove All"
         End
         Begin VB.Menu mnuInfo 
            Caption         =   "Info"
         End
      End
      Begin VB.Menu mnuColours 
         Caption         =   "Colours"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFileName As String ' String of file to open
Dim strText As String ' contents of file
Dim strFilter As String ' common dialog filter string
Dim strBuffer As String ' string buffer variable
Dim Filehandle ' variable to hold file handle

Dim answer

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'RedHoverColour cmdRemove
End Sub

Private Sub Check2_Click()
NotOntop frmMain
answer = MsgBox("this feature is currently unnavailable", vbOKOnly + vbCritical, "mp3 d00d")
If answer = vbOK Then
Ontop frmMain
End If
End Sub

Public Sub chkPlayList_Click()
If chkPlayList.Value = 1 Then
frmPlayList.Show
frmPlayList.Top = frmMain.Top
frmPlayList.Left = frmMain.Left + frmMain.Width
'If lstfiles.ListCount >= 1 Then
'OpenFileList2 App.path + "\temp.mpd"
'OpenMainFileList2 App.path + "\temp2.mpd"
'Else: frmPlayList.lstfiles.List = ""
'End If
Ontop frmPlayList

ElseIf chkPlayList.Value = 0 Then
frmPlayList.Visible = False
End If

End Sub


Private Sub cmdNext_Click()

On Error GoTo bug
Dim total As Integer
total = lstfiles.ListCount
If optShuffle.Value = True Then
        Randomize
        Dim num As Integer
        num = Int((lstfiles.ListCount * Rnd))
        
        lstfiles.Selected(num) = True
        lstNames.Selected(num) = True
        frmPlayList.lstfiles.Selected(num) = True
        frmPlayList.lstNames.Selected(num) = True
        
        MP.FileName = lstfiles.Text
        MP.Play
        tmrScroll.Enabled = True
        SldProgress.Max = MP.Duration
        lbltitle.Caption = lstNames.Text
        
Else
If lstfiles.ListIndex = total - 1 Then
    lstfiles.ListIndex = 0
    lstNames.ListIndex = 0
    txtPath = lstfiles.Text
    MP.FileName = txtPath
    MP.Play
    SldProgress.Max = MP.Duration
Else

lstfiles.ListIndex = lstfiles.ListIndex + 1
lstNames.ListIndex = lstNames.ListIndex + 1

txtPath.Text = lstfiles.Text

MP.FileName = txtPath.Text
lbltitle.Caption = lstNames.Text

MP.Play
SldProgress.Max = MP.Duration

End If
bug:
    If Err.number = 380 Then
        NotOntop frmMain

        answer = MsgBox("There is no song in the playlist to goto", vbExclamation, "mp3 d00d")
            If answer = vbOK Then
                Ontop frmMain
            End If
    End If
End If
End Sub

Private Sub cmdPause_Click()
' this is the even for if we want to pause the song

'set and use the If Else statement. If the label has Pause as its caption then
'pause the mp3
'set the caption to Resume
'disable the play label
If MP.PlayState = mpPaused Then
MP.Play

cmdPlay.Enabled = True

'if the caption says resume then
'play the mp3 from its current point in the song
'set the caption to pause
'allow the play label to be enabled
Else
MP.Pause
'lblPause.Caption = "Pause"
cmdPlay.Enabled = False
End If

End Sub

Private Sub cmdPlay_Click()
On Error GoTo bug

    MP.FileName = txtPath.Text
    'play the selected mp3

MP.Play
'set the sliders(progress bar) maxamum number (the furthest it will go) to the duration
' of the song
SldProgress.Max = MP.Duration
'enable the timer so that it can start to count the time used in the song
tmrTime.Enabled = True
'MP.Play
'allow the pause label to become enabled
cmdPause.Enabled = True

tmrScroll.Enabled = True
'End If
lbltitle.Caption = lstNames.Text
bug:
'If Err.number = 63 Then
'    NotOntop frmMain
'    answer = MsgBox("Please select a song to play.  If there is no song in the playlist please add songs by selecting the add menu", vbExclamation, "Mp3 d00d")
'    If answer = vbOK Then
'Ontop frmMain
'End If
'End If

End Sub

Private Sub cmdPrev_Click()
On Error Resume Next

'this event is for when the use presses the next label, it skips to the next song:


'this sets the listindex for the list so that it grabs the previous song in the list and
'not the current one
If lstfiles.ListIndex = 0 Then
    frmPlayList.lstNames.ListIndex = frmPlayList.lstNames.ListCount - 1
    frmPlayList.lstfiles.ListIndex = frmPlayList.lstfiles.ListCount - 1
    frmMain.lstNames.ListIndex = frmMain.lstNames.ListCount - 1
    frmMain.lstfiles.ListIndex = frmMain.lstfiles.ListCount - 1

    MP.FileName = txtPath
    MP.Play
    tmrScroll.Enabled = True
Else
lstfiles.ListIndex = lstfiles.ListIndex - 1
lstNames.ListIndex = lstNames.ListIndex - 1
SldProgress.Max = MP.Duration

'this will set the path for the mp3 so that it can be played using the media control
txtPath.Text = lstfiles.Text

'sets the path for the media player
MP.FileName = txtPath.Text

SldProgress.Max = MP.Duration
'plays selected mp3
MP.Play
lbltitle.Caption = lstNames.Text
bug:
    If Err.number = 381 Then
    NotOntop frmMain
      answer = MsgBox("There is no song in the playlist to goto", vbExclamation, "mp3 d00d")
      If answer = vbOK Then
Ontop frmMain
End If
    End If
End If
End Sub

Private Sub cmdStop_Click()
MP.Stop
cmdPause.Enabled = False
cmdPlay.Enabled = True

End Sub

Private Sub Command1_Click()
Unload Me
End
End Sub

Private Sub Command2_Click()
frmMain.WindowState = 1
frmPlayList.Visible = False
frmMain.Caption = lstNames.Text
End Sub

Private Sub Command3_Click()
NotOntop frmMain
answer = MsgBox("Later", vbExclamation, "mp3 d00d")
If answer = vbOK Then
Ontop frmMain
End If
End Sub


Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
'when clicked this shows the form addfile, so that the user can add their mp3's
frmAddFile.Show


End Sub

Private Sub Form_Load()
   'Syntax - Associate(AppTitle, FileExtension, FileType, IconFileName, Parameters)
    '
    'AppTitle - The key that will be created in classes root directory (e.g. "MyApp"
    'FileExtension - The extension of the file to associate (e.g. ".EXT")
    'FileType - The file type of the application, that will be used as a short description (e.g. "Extension Text")
    'IconFileName - Specifies the path of the icon to be used in the application (e.g. "C:\MyApp\Icon.icon"), can be also icon in a libary (e.g. "C:\MyApp\MyDll.dll,2")
    'Parameters - Any parameters that might be used for the applicationf ile (e.g. "/parameter" or "/param1 /param2" etc..)
    'Associate "Project1", ".mpd", "mp3 d00d Playlist", "Shell32.dll,25"
    Associate "mp3 d00d", ".mpd", "mp3 d00d Playlist", "Shell32.dll,25"

If Command <> "" Then 'file found
'      MsgBox "The File " & Command & " was loaded!", vbInformation
'    End If
    frmMain.Picture = LoadPicture(App.path + "\pics\header.jpg")
    lstfiles.Clear
    frmPlayList.lstfiles.Clear
    frmPlayList.lstNames.Clear
    lstNames.Clear
    OpenPathList Command
    OpenFileList "C:\windows\" & GetFileTitle(Command)
    'loadMainForm
    LoadMPDForm


Else
frmMain.Picture = LoadPicture(App.path + "\pics\header.jpg")
LoadMainForm
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuFile
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
MoveForm frmMain
End If
If chkPlayList.Value = 1 Then
frmPlayList.Top = frmMain.Top
frmPlayList.Left = frmMain.Left + frmMain.Width
End If
End Sub


Private Sub Form_Resize()
'frmMain.Height = 2235

If frmMain.WindowState = 0 And chkPlayList.Value = vbChecked Then
    frmPlayList.Visible = True
    frmMain.Caption = ""
    frmMain.Height = 2235
ElseIf frmMain.WindowState = 0 Then
    frmMain.Caption = ""
    frmMain.Height = 2235
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
SavePathList "c:\windows\temp.mpd"
SaveMainFileList "C:\windows\temp2.mpd"
If frmPlayList.Visible = True Then
    SaveSetting "mp3 d00d", "Properties", "Playlist", "Open"
Else
    SaveSetting "mp3 d00d", "Properties", "Playlist", "Close"
End If
End Sub


Private Sub lstfiles_Click()
txtPath = lstfiles.Text
End Sub

Private Sub lstfiles_DblClick()
'lblPlay_Click
End Sub

Private Sub lstfiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuFile
End If
End Sub

Private Sub lstNames_Click()


If lstNames.Selected(lstNames.ListIndex) = True Then
lstfiles.Selected(lstNames.ListIndex) = True
frmPlayList.lstfiles.Selected(lstNames.ListIndex) = True
frmPlayList.lstNames.Selected(lstNames.ListIndex) = True
txtPath = lstfiles.Text
Else
lstNames.Selected(lstNames.ListIndex) = False
lstfiles.Selected(lstfiles.ListIndex) = False
frmPlayList.lstNames.Selected(lstNames.ListIndex) = False
frmPlayList.lstfiles.Selected(lstfiles.ListIndex) = False
End If
'bug:
    'If Err.number = "381" Then
       ' End If
End Sub

Private Sub lstNames_DblClick()
cmdPlay_Click
End Sub

Private Sub mnuColours_Click()
NotOntop frmMain
frmColours.Show
End Sub

Private Sub mnuNext_Click()
cmdNext_Click
End Sub

Private Sub mnuPause_Click()
cmdPause_Click
End Sub

Private Sub mnuPlay_Click()
cmdPlay_Click
End Sub

Private Sub mnuPrevious_Click()
cmdPrev_Click
End Sub

Private Sub mnuRemoveAll_Click()
'cmdRemoveAll_Click
End Sub



Private Sub mnuStop_Click()
cmdStop_Click
End Sub




Private Sub SldProgress_scroll()
MP.CurrentPosition = SldProgress.Value
End Sub

Private Sub sldVolume_scroll()
'this is used for when the use wants to adjust the volume

'dimmension, or declare the variable MaxVol to be used as a variable
Dim MaxVol
' same as above

Dim SetMin As Integer, SetVal As Integer
'sets the variable MaxVol to the sliders value. this means that its value is -2500

MaxVol = sldVolume.Value - 2500
'this sets the volume on the mediaplayer control to the variable MaxVol

MP.Volume = MaxVol
'if an error occurs, then goto BIGERROR

On Error GoTo BIGERROR
'set the set min variable to the sliders minimum current setting

SetMin = sldVolume.min
'sets the new value to the slider

SetVal = sldVolume.Value
'come here on err

BIGERROR:
'exists this funciton
Exit Sub

End Sub
Private Sub MP_EndOfStream(ByVal Result As Long)
 'make mp3 d00d play the next song in the list
 Dim totcount As Variant
 Dim totall As Integer
 totall = lstfiles.ListCount
 totcount = lstfiles.ListCount - 1
 Dim number, jared
 jared = lstfiles.ListIndex = lstfiles.ListCount - 1
 
 If optRepeat.Value = True And lstfiles.ListIndex = totcount Then
        lstfiles.ListIndex = 0
        lstNames.ListIndex = 0
        MP.FileName = lstfiles.Text
        SldProgress.Max = MP.Duration
        lbltitle.Caption = lstNames.Text
        MP.Play
         If frmMain.WindowState = 1 Then
                frmMain.Caption = frmMain.lstNames.Text
            End If
ElseIf optRepeat.Value = True Then
        lstfiles.ListIndex = lstfiles.ListIndex + 1
        lstNames.ListIndex = lstNames.ListIndex + 1
        MP.FileName = lstfiles.Text
        SldProgress.Max = MP.Duration
        lbltitle.Caption = lstNames.Text
        MP.Play
            If frmMain.WindowState = 1 Then
                frmMain.Caption = frmMain.lstNames.Text
            End If
ElseIf optNormal.Value = True And lstfiles.ListIndex = totcount Then
    MP.Stop
ElseIf optShuffle.Value = True Then
        Randomize
        Dim num As Integer
        num = Int((lstfiles.ListCount * Rnd))
        
        lstfiles.Selected(num) = True
        lstNames.Selected(num) = True
        frmPlayList.lstfiles.Selected(num) = True
        frmPlayList.lstNames.Selected(num) = True
        
        MP.FileName = lstfiles.Text
        SldProgress.Max = MP.Duration
        lbltitle.Caption = lstNames.Text
        MP.Play
            If frmMain.WindowState = 1 Then
                frmMain.Caption = frmMain.lstNames.Text
            End If
Else
        lstfiles.ListIndex = lstfiles.ListIndex + 1
        lstNames.ListIndex = lstNames.ListIndex + 1
        MP.FileName = lstfiles.Text
        SldProgress.Max = MP.Duration
        lbltitle.Caption = frmMain.lstNames.Text
        MP.Play
            If frmMain.WindowState = 1 Then
                frmMain.Caption = frmMain.lstNames.Text
            End If
End If

End Sub



Private Sub tmrOpen_Timer()
'OpenFileList "C:\windows\" & lblOpen.Caption
'tmrOpen.Enabled = False

End Sub

Private Sub tmrSave_Timer()
SaveFileList "C:\windows\" & lblSave.Caption
tmrSave.Enabled = False

End Sub

Private Sub tmrScroll_Timer()
  If lbltitle.Left < -1000 Then
        lbltitle.Left = 7000
    Else
        lbltitle.Left = Val(lbltitle.Left) - 40
    End If
End Sub

Private Sub tmrTime_Timer()
SldProgress.Value = MP.CurrentPosition
Dim mp3Time
mp3Time = MP.CurrentPosition
Dim min, sec As Integer
min = mp3Time \ 60
sec = mp3Time - (min * 60)
If sec = "-1" Then sec = "0"
lblTime.Caption = min & ":" & sec
End Sub

Public Sub SaveFileList(path As String)
    
        
On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open path For Output As #1

    
    For Save = 0 To lstNames.ListCount - 1
        Print #1, lstNames.List(Save)
    Next Save
    Close #1
End Sub
Public Sub OpenPathList(path As String)
      Filehandle = FreeFile
    Dim strJared As String
 Open path For Input As #Filehandle

       On Error GoTo shit2
shit:
       Input #Filehandle, strJared
       
      lstfiles.AddItem strJared
      frmPlayList.lstfiles.AddItem strJared
      GoTo shit
shit2:
Close #Filehandle
'lstfiles.RemoveItem lstfiles.ListIndex = 0
End Sub

Public Sub OpenMainFileList(path As String)
      Filehandle = FreeFile
    Dim strJared As String
 Open path For Input As #Filehandle

       On Error GoTo shit2
shit:
       Input #Filehandle, strJared
       
      lstNames.AddItem strJared
       
      GoTo shit
shit2:
Close #Filehandle
'lstfiles.RemoveItem lstfiles.ListIndex = 0
End Sub
Public Sub SaveMainFileList(path As String)
    
        
On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open path For Output As fFile

    
    For Save = 0 To lstNames.ListCount - 1
        Print #fFile, lstNames.List(Save)
    Next Save
    Close fFile
End Sub
Public Sub OpenMainFileList2(path As String)
      Filehandle = FreeFile
    Dim strJared As String
 Open path For Input As #Filehandle

       On Error GoTo shit2
shit:
       Input #Filehandle, strJared
       
      frmPlayList.lstNames.AddItem strJared
       
      GoTo shit
shit2:
Close #Filehandle
'lstfiles.RemoveItem lstfiles.ListIndex = 0
End Sub
Public Sub OpenFileList2(path As String)
      Filehandle = FreeFile
    Dim strJared As String
 Open path For Input As #Filehandle

       On Error GoTo shit2
shit:
       Input #Filehandle, strJared
       
      frmPlayList.lstfiles.AddItem strJared
       
      GoTo shit
shit2:
Close #Filehandle
'lstfiles.RemoveItem lstfiles.ListIndex = 0
End Sub
Public Sub SavePathList(path As String)
    
        
On Error Resume Next
    Dim Save As Long
    Dim fFile As Integer
    fFile = FreeFile
    Open path For Output As #1

    
    For Save = 0 To lstfiles.ListCount - 1
        Print #1, lstfiles.List(Save)
    Next Save
    Close #1
End Sub

Public Sub OpenFileList(path As String)
      Filehandle = FreeFile
    Dim strJared As String
 Open path For Input As #Filehandle

       On Error GoTo shit2
shit:
       Input #Filehandle, strJared
       
      lstNames.AddItem strJared
      frmPlayList.lstNames.AddItem strJared
       
      GoTo shit
shit2:
Close #Filehandle
'lstfiles.RemoveItem lstfiles.ListIndex = 0
End Sub





