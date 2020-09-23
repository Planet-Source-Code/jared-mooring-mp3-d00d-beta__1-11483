VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColours 
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optCheckBack 
      Caption         =   "Check box Background"
      Height          =   375
      Left            =   360
      TabIndex        =   33
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   4215
      Left            =   2520
      TabIndex        =   24
      Top             =   480
      Width           =   3135
      Begin VB.CommandButton cmdChange 
         Caption         =   "Select Colour"
         Height          =   375
         Left            =   600
         TabIndex        =   31
         Top             =   3720
         Width           =   2175
      End
      Begin VB.DriveListBox Drive 
         Height          =   315
         Left            =   480
         TabIndex        =   30
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton Button 
         Caption         =   "Command1"
         Height          =   495
         Left            =   960
         TabIndex        =   28
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox List 
         Height          =   2400
         Left            =   840
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox CheckBox 
         Caption         =   "Check1"
         Height          =   495
         Left            =   1080
         TabIndex        =   26
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox Background 
         Height          =   2055
         Left            =   720
         ScaleHeight     =   1995
         ScaleWidth      =   1755
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   1215
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label LABEL 
         Caption         =   "Label1"
         Height          =   495
         Left            =   1080
         TabIndex        =   29
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.OptionButton optPlaylistButtonDown 
      Caption         =   "Play list move down Button"
      Height          =   375
      Left            =   5640
      TabIndex        =   23
      Top             =   2640
      Width           =   2295
   End
   Begin VB.OptionButton optPlayButtonUp 
      Caption         =   "Play List Move Up Button"
      Height          =   375
      Left            =   5640
      TabIndex        =   22
      Top             =   2280
      Width           =   2415
   End
   Begin VB.OptionButton optHeaderBackground 
      Caption         =   "Header Background"
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   5280
      Width           =   1935
   End
   Begin VB.OptionButton OptPause 
      Caption         =   "Pause Song"
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton optNext 
      Caption         =   "Next song"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.OptionButton optStop 
      Caption         =   "Stop Song"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton optCheckBox 
      Caption         =   "Check box Text"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   4560
      Width           =   1575
   End
   Begin VB.OptionButton optBackground 
      Caption         =   "Background"
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton optPlay 
      Caption         =   "Play Song"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton optButtonHover 
      Caption         =   "Button Hover Colour"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
   End
   Begin VB.OptionButton optListBack 
      Caption         =   "List Background"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   2760
      Width           =   1695
   End
   Begin VB.OptionButton optDriveListText 
      Caption         =   "Drive List Text"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3840
      Width           =   1695
   End
   Begin VB.OptionButton optListText 
      Caption         =   "List Text"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.OptionButton optButton 
      Caption         =   "Buttons"
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton OptPlayListButton 
      Caption         =   "Play List Button"
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.OptionButton optPlayListHover 
      Caption         =   "Playlist Hover Colour"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.OptionButton optPrev 
      Caption         =   "Previous Song"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton optTitleText 
      Caption         =   "Title Label Text"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   3600
      Width           =   1455
   End
   Begin VB.OptionButton optTimeText 
      Caption         =   "Time Label Text"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.OptionButton optTimeBack 
      Caption         =   "Time Label Background"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   3960
      Width           =   2175
   End
   Begin VB.OptionButton optScrollBack 
      Caption         =   "Info Label (scrolling label) Background"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   5280
      Width           =   3015
   End
   Begin VB.OptionButton optScrollText 
      Caption         =   "Info Label (scrolling label) Text"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   5520
      Width           =   2775
   End
   Begin VB.OptionButton optDriveListBack 
      Caption         =   "Drive List Background"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog cdColour 
      Left            =   6240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton optTitleBack 
      Caption         =   "Title Label Background"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   3480
      TabIndex        =   32
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   8040
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line2 
      X1              =   5760
      X2              =   8160
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2280
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "frmColours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intCaption As Integer


Private Sub cmdChange_Click()

cdColour.Flags = cdlCCRGBInit
cdColour.ShowColor

intCaption = Label1.Caption

Select Case intCaption
    
    Case 1
        WriteINI App.path + "\colour.dat", "Label_Prev", "Colour", cdColour.Color
    Case 2
        WriteINI App.path + "\colour.dat", "Label_Play", "Colour", cdColour.Color
    Case 3
        WriteINI App.path + "\colour.dat", "Label_Pause", "Colour", cdColour.Color
    Case 4
        WriteINI App.path + "\colour.dat", "Label_Stop", "Colour", cdColour.Color
    Case 5
        WriteINI App.path + "\colour.dat", "Label_Next", "Colour", cdColour.Color
    Case 6
        WriteINI App.path + "\colour.dat", "List_Background", "Colour", cdColour.Color
    Case 7
        WriteINI App.path + "\colour.dat", "List_Text", "Colour", cdColour.Color
    Case 8
        WriteINI App.path + "\colour.dat", "drive_back", "Colour", cdColour.Color
    Case 9
        WriteINI App.path + "\colour.dat", "drive_text", "Colour", cdColour.Color
    Case 10
        WriteINI App.path + "\colour.dat", "check_back", "Colour", cdColour.Color
    Case 11
        WriteINI App.path + "\colour.dat", "check_text", "Colour", cdColour.Color
    Case 12
        WriteINI App.path + "\colour.dat", "Background", "Colour", cdColour.Color
    Case 13
        WriteINI App.path + "\colour.dat", "Buttons", "Colour", cdColour.Color
    Case 14
        WriteINI App.path + "\colour.dat", "Button_Hover", "Colour", cdColour.Color
    Case 15
        WriteINI App.path + "\colour.dat", "PlayList_Button", "Colour", cdColour.Color
    Case 16
        WriteINI App.path + "\colour.dat", "Playlist_Hover", "Colour", cdColour.Color
    Case 17
        WriteINI App.path + "\colour.dat", "Move_up", "Colour", cdColour.Color
    Case 18
        WriteINI App.path + "\colour.dat", "Move_Down", "Colour", cdColour.Color
    Case 19
        WriteINI App.path + "\colour.dat", "Title_Back", "Colour", cdColour.Color
    Case 20
        WriteINI App.path + "\colour.dat", "Title_text", "Colour", cdColour.Color
    Case 21
        WriteINI App.path + "\colour.dat", "Time_Back", "Colour", cdColour.Color
    Case 22
        WriteINI App.path + "\colour.dat", "Time_Text", "Colour", cdColour.Color
    Case 23
        WriteINI App.path + "\colour.dat", "Header_Shape", "Colour", cdColour.Color
    Case 24
        WriteINI App.path + "\colour.dat", "Info_Scroller_Text", "Colour", cdColour.Color
    Case 25
        WriteINI App.path + "\colour.dat", "Info_Scroller_Text", "Colour", cdColour.Color
    
End Select

    
End Sub

Private Sub optBackground_Click()
BackgroundClicked
Background.Visible = True
Label1.Caption = "12"

End Sub

Private Sub optBorder_Click()

End Sub

Private Sub optButton_Click()
ButtonClicked

Button.Visible = True
Button.Caption = "Regular Button"
Label1.Caption = "13"
End Sub

Private Sub optButtonHover_Click()
ButtonClicked

Button.Visible = True
Button.Caption = "Regular Button"
Label1.Caption = "14"
End Sub

Private Sub optCheckBack_Click()
CheckClicked
CheckBox.Visible = True
CheckBox.Caption = "Check Box"
Label1.Caption = "10"
End Sub

Private Sub optCheckBox_Click()
CheckClicked

CheckBox.Visible = True
CheckBox.Caption = "Check Box"

Label1.Caption = "11"
End Sub

Private Sub optDriveListBack_Click()
DriveClicked

Drive.Visible = True
Label1.Caption = "8"
End Sub

Private Sub optDriveListText_Click()
DriveClicked

Drive.Visible = True
Label1.Caption = "9"
End Sub

Private Sub optHEaderBackground_Click()
ShapeClicked

Label1.Caption = "23"
Shape1.Visible = True

End Sub

Private Sub optListBack_Click()
ListClicked

List.Visible = True
List.AddItem "ListBox"

Label1.Caption = "6"

End Sub

Private Sub optListText_Click()
ListClicked

List.Visible = True
List.AddItem "List Box"
Label1.Caption = "7"
End Sub

Private Sub optNext_Click()
LabelClicked
LABEL.Visible = True
LABEL.Caption = "Next"

Label1.Caption = "5"

End Sub

Private Sub OptPause_Click()
LabelClicked

LABEL.Visible = True
LABEL.Caption = "Pause"

Label1.Caption = "3"


End Sub

Private Sub optPlay_Click()
LabelClicked
LABEL.Visible = True
LABEL.Caption = "Play"

Label1.Caption = "2"
End Sub

Private Sub optPlayButtonUp_Click()
ButtonClicked
Button.Visible = True
Button.Caption = "Playlist Up Button"
Label1.Caption = "17"
End Sub

Private Sub OptPlayListButton_Click()
ButtonClicked

Button.Visible = True
Button.Caption = "Playlist Button"

Label1.Caption = "15"
End Sub

Private Sub optPlaylistButtonDown_Click()
ButtonClicked
Button.Visible = True
Button.Caption = "PlayList Down Button"
Label1.Caption = "18"
End Sub

Private Sub optPlayListHover_Click()
ButtonClicked
Button.Visible = True
Button.Caption = "Playlist Hover Colour"
Label1.Caption = "16"
End Sub

Private Sub optPrev_Click()
LabelClicked

LABEL.Visible = True
LABEL.Caption = "Previous"

Label1.Caption = "1"
End Sub

Private Sub optScrollBack_Click()
LabelClicked

LABEL.Visible = True
LABEL.Caption = "Scrolling Background"

Label1.Caption = "25"
End Sub

Private Sub optScrollText_Click()
LabelClicked

LABEL.Visible = True
LABEL.Caption = "Scrolling Text"

Label1.Caption = "24"

End Sub

Private Sub optStop_Click()
LabelClicked

LABEL.Visible = True
LABEL.Caption = "Stop"

Label1.Caption = "4"
End Sub

Private Sub optTimeBack_Click()
LabelClicked

LABEL.Visible = True
LABEL.Caption = "Time Background"
Label1.Caption = "21"
End Sub

Private Sub optTimeText_Click()
LabelClicked

LABEL.Visible = True
LABEL.Caption = "Time Text"
Label1.Caption = "22"
End Sub

Private Sub optTitleBack_Click()
LabelClicked

LABEL.Visible = True
LABEL.Caption = "Title Backgound"
Label1.Caption = "19"
End Sub

Private Sub optTitleText_Click()
LabelClicked

LABEL.Visible = True
LABEL.Caption = "Title Text"
Label1.Caption = "20"
End Sub


Public Sub LabelClicked()
Drive.Visible = False
List.Visible = False
Background.Visible = False
CheckBox.Visible = False
Button.Visible = False
Shape1.Visible = False
End Sub
Public Sub DriveClicked()
List.Visible = False
LABEL.Visible = False
Background.Visible = False
CheckBox.Visible = False
Button.Visible = False
Shape1.Visible = False
End Sub
Public Sub ListClicked()
Button.Visible = False
Drive.Visible = False
LABEL.Visible = False
Background.Visible = False
CheckBox.Visible = False
Shape1.Visible = False
End Sub
Public Sub ShapeClicked()
Button.Visible = False
Drive.Visible = False
List.Visible = False
LABEL.Visible = False
Background.Visible = False
CheckBox.Visible = False
End Sub
Public Sub CheckClicked()
Drive.Visible = False
List.Visible = False
LABEL.Visible = False
Background.Visible = False
Button.Visible = False
Shape1.Visible = False
End Sub
Public Sub ButtonClicked()
Drive.Visible = False
List.Visible = False
LABEL.Visible = False
Background.Visible = False
CheckBox.Visible = False
Shape1.Visible = False
End Sub
Public Sub BackgroundClicked()
Drive.Visible = False
List.Visible = False
LABEL.Visible = False
CheckBox.Visible = False
Shape1.Visible = False
Button.Visible = False
End Sub
