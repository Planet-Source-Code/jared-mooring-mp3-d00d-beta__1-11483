VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmold 
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   498
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   1455
      Left            =   5400
      TabIndex        =   10
      ToolTipText     =   "Volume"
      Top             =   960
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   2566
      _Version        =   393216
      Orientation     =   1
      Max             =   2500
      TickStyle       =   3
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin MSComctlLib.Slider SldProgress 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   10
      SmallChange     =   0
      TickStyle       =   2
      TickFrequency   =   0
      TextPosition    =   1
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   6600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   480
      TabIndex        =   1
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label lblPause 
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblNext 
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblPlay 
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblStop 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblPrev 
      BackStyle       =   0  'Transparent
      Caption         =   "Prev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin MediaPlayerCtl.MediaPlayer MP 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
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
End
Attribute VB_Name = "Frmold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
frmMain.Show
End Sub

Private Sub Form_Load()
List1.AddItem "C:\Mp3's\Murphy, Eddie - Raw.mp3"
List1.AddItem "C:\Mp3's\Misc\Dune- Star Child.mp3"
List1.AddItem "C:\Mp3's\Misc\Slipknot-02(sic).mp3"
List1.AddItem "C:\Mp3's\Misc\Delerium - Wisdom.mp3"
List1.AddItem "C:\Mp3's\Misc\Paffendorf - Smile.mp3"
List1.AddItem "C:\Mp3's\Misc\Stardust-Music_sounds_better_with_you.mp3"
List1.AddItem "C:\Mp3's\Misc\mb_bent.mp3"

'frmMain.Picture = LoadPicture(App.Path + "\pics\back.jpg")
End Sub



Private Sub lblNext_Click()
List1.ListIndex = List1.ListIndex + 1
        Text1.Text = List1.Text
        MP.FileName = Text1.Text
        MP.Play
End Sub

Private Sub lblPause_Click()
If lblPause.Caption = "Pause" Then
MP.Pause
lblPause.Caption = "Resume"
lblPlay.Enabled = False
Else
MP.Play
lblPause.Caption = "Pause"
lblPlay.Enabled = True
End If


End Sub

Private Sub lblPlay_Click()
MP.FileName = Text1
MP.Play
SldProgress.Max = MP.Duration
Timer1.Enabled = True
MP.Play
lblPause.Enabled = True
End Sub

Private Sub lblPrev_Click()
List1.ListIndex = List1.ListIndex - 1
        Text1.Text = List1.Text
        MP.FileName = Text1.Text
        MP.Play
End Sub

Private Sub lblStop_Click()
MP.Stop
End Sub

Private Sub List1_Click()
Text1.Text = List1.Text
End Sub

Private Sub List1_dblClick()
MP.FileName = Text1
Text1 = List1.Text
MP.FileName = Text1
MP.Play
SldProgress.Max = MP.Duration
Timer1.Enabled = True
lblPause.Enabled = True

End Sub


Private Sub SldProgress_scroll()
MP.CurrentPosition = SldProgress.Value
End Sub


Private Sub sldVolume_scroll()
Dim MaxVol
Dim SetMin As Integer, SetVal As Integer
MaxVol = sldVolume.Value - 2500
MP.Volume = MaxVol
On Error GoTo BIGERROR
SetMin = sldVolume.min
SetVal = sldVolume.Value
BIGERROR:
Exit Sub
End Sub
Private Sub Timer1_Timer()
SldProgress.Value = MP.CurrentPosition
Dim mp3Time
mp3Time = MP.CurrentPosition
Dim min, sec As Integer
min = mp3Time \ 60
sec = mp3Time - (min * 60)
If sec = "-1" Then sec = "0"
lblTime.Caption = min & ":" & sec
End Sub

Private Sub MP_EndOfStream(ByVal Result As Long)
 Dim totcount As Variant
 totcount = List1.ListCount - 1
 If totcount = List1.ListCount Then
 
 Else
On Error Resume Next
 List1.ListIndex = List1.ListIndex + 1
MP.FileName = List1.Text
 MP.Play
 End If

End Sub
