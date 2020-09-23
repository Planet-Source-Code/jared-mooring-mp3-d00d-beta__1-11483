Attribute VB_Name = "basSubs"
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE


Global Const conHwndTopmost = -1
Global Const conSwpNoActivate = &H10
Global Const conSwpShowWindow = &H40

Public Const SPI_SETDESKWALLPAPER = 20
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Public Sub WhiteHoverColour(LABEL As LABEL)
LABEL.ForeColor = vbWhite
End Sub
Public Sub OtherHoverColour(LABEL As LABEL)
LABEL.ForeColor = &HFFFF00
End Sub
Public Sub RedHoverColour(Button As CommandButton)
Button.BackColor = vbRed
End Sub
Public Sub CyanHoverColour(Button As CommandButton)
Button.BackColor = vbCyan
End Sub

'code borrowed from planetsourcecode.com
Function MoveForm(mHwnd As Form)
ReleaseCapture
SendMessage mHwnd.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 1
End Function

Public Sub NotOntop(FormName As Form)
'Make a form not always ontop of other windows
On Error GoTo error
Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
Public Sub Ontop(FormName As Form)
'Make a form always ontop of other windows
On Error GoTo error
Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
Public Function PlayListPathExist() As Boolean

If Dir("c:\windows\temp.mpd") = "" Then
    PlayListPathExist = False
    Exit Function
End If
 PlayListPathExist = True
End Function
Public Function PlayListFileExist() As Boolean

If Dir(App.path + "\temp2.mpd") = "" Then
    PlayListFileExist = False
    Exit Function
End If
 PlayListFileExist = True
End Function

Public Sub LoadMainColours()
'frmMain.lblPrev.ForeColor = ReadINI(App.path + "\colour.dat", "Label_Prev", "Colour")
'frmMain.lblPlay.ForeColor = ReadINI(App.path + "\colour.dat", "Label_Play", "Colour")
'frmMain.lblPause.ForeColor = ReadINI(App.path + "\colour.dat", "Label_Pause", "Colour")
'frmMain.lblStop.ForeColor = ReadINI(App.path + "\colour.dat", "Label_Stop", "Colour")
'frmMain.lblNext.ForeColor = ReadINI(App.path + "\colour.dat", "Label_Next", "Colour")
'frmMain.lstNames.BackColor = ReadINI(App.path + "\colour.dat", "List_Background", "Colour")
'frmMain.lstNames.ForeColor = ReadINI(App.path + "\colour.dat", "List_Text", "Colour")
'frmMain.lblPrev.BackColor = ReadINI(App.path + "\colour.dat", "drive_back", "Colour")

'frmMain.lblPrev.BackColor = ReadINI(App.path + "\colour.dat", "drive_text", "Colour")
'frmMain.Check1.BackColor = ReadINI(App.path + "\colour.dat", "check_back", "Colour")
'frmMain.Check1.ForeColor = ReadINI(App.path + "\colour.dat", "check_text", "Colour")
'frmMain.Check2.BackColor = ReadINI(App.path + "\colour.dat", "check_back", "Colour")
'frmMain.Check2.ForeColor = ReadINI(App.path + "\colour.dat", "check_text", "Colour")
'frmMain.BackColor = ReadINI(App.path + "\colour.dat", "Background", "Colour")
''frmMain.lblPrev.BackColor = ReadINI(App.path + "\colour.dat", "Buttons", "Colour")
''frmMain.lblPrev.BackColor = ReadINI(App.path + "\colour.dat", "Button_Hover", "Colour")
'
'frmMain.cmdRemove.BackColor = ReadINI(App.path + "\colour.dat", "PlayList_Button", "Colour")
'frmMain.cmdRemoveAll.BackColor = ReadINI(App.path + "\colour.dat", "PlayList_Button", "Colour")
'frmMain.cmdRemoveFile.BackColor = ReadINI(App.path + "\colour.dat", "PlayList_Button", "Colour")
'frmMain.cmdAdd.BackColor = ReadINI(App.path + "\colour.dat", "PlayList_Button", "Colour")
'frmMain.cmdAddFile.BackColor = ReadINI(App.path + "\colour.dat", "PlayList_Button", "Colour")
'frmMain.cmdAddDir.BackColor = ReadINI(App.path + "\colour.dat", "PlayList_Button", "Colour")
'frmMain.cmdList.BackColor = ReadINI(App.path + "\colour.dat", "PlayList_Button", "Colour")
'frmMain.cmdListOpen.BackColor = ReadINI(App.path + "\colour.dat", "PlayList_Button", "Colour")
'frmMain.cmdListSave.BackColor = ReadINI(App.path + "\colour.dat", "PlayList_Button", "Colour")
'
'frmMain.cmdMoveUp.BackColor = ReadINI(App.path + "\colour.dat", "Move_up", "Colour")
'frmMain.cmdMoveDown.BackColor = ReadINI(App.path + "\colour.dat", "Move_Down", "Colour")
'frmMain.lbltitle.BackColor = ReadINI(App.path + "\colour.dat", "Title_Back", "Colour")
'frmMain.Frame1.BackColor = ReadINI(App.path + "\colour.dat", "Title_Back", "Colour")
'frmMain.lbltitle.ForeColor = ReadINI(App.path + "\colour.dat", "Title_text", "Colour")
'frmMain.lblTime.BackColor = ReadINI(App.path + "\colour.dat", "Time_Back", "Colour")
'frmMain.lblTime.ForeColor = ReadINI(App.path + "\colour.dat", "Time_Text", "Colour")
'frmMain.lblPrev.BackColor = ReadINI(App.path + "\colour.dat", "Header_Shape", "Colour")
'frmMain.lblPrev.BackColor = ReadINI(App.path + "\colour.dat", "Info_Scroller_Text", "Colour")
'frmMain.lblPrev.BackColor = ReadINI(App.path + "\colour.dat", "Info_Scroller_Text", "Colour")
'frmMain.Shape2.BackColor = ReadINI(App.path + "\colour.dat", "Header_Shape", "Colour")
End Sub
Public Sub LoadMainForm()
'basSubs.LoadMainColours
Ontop frmMain
frmMain.sldVolume.Value = "2500"
'SaveList
If PlayListPathExist = True Then
frmMain.OpenPathList "C:\windows\temp.mpd"
frmMain.OpenFileList2 "C:\windows\temp.mpd"

'End If
'If PlayListFileExist = True Then
frmMain.OpenMainFileList "C:\windows\temp2.mpd"
frmMain.OpenMainFileList2 "C:\windows\temp2.mpd"
End If
frmMain.optNormal.Value = True
'frmPlayList.Show
If frmMain.lstfiles.ListCount > 0 Then
    frmMain.lstfiles.Selected(0) = True
    frmMain.lstNames.Selected(0) = True
    frmPlayList.lstNames.Selected(0) = True
    frmPlayList.lstfiles.Selected(0) = True
End If
frmMain.lblPlaylist.Caption = GetSetting("mp3 d00d", "properties", "Playlist")
    If frmMain.lblPlaylist.Caption = "Close" Then
        frmPlayList.Visible = False
    Else
       frmMain.chkPlayList.Value = 1
       frmMain.chkPlayList_Click
    End If
End Sub
Public Sub LoadMPDForm()
'basSubs.LoadMainColours
Ontop frmMain
frmMain.sldVolume.Value = "2500"


frmMain.optNormal.Value = True
'frmPlayList.Show
If frmMain.lstfiles.ListCount > 0 Then
    frmMain.lstfiles.Selected(0) = True
    frmMain.lstNames.Selected(0) = True
    frmPlayList.lstNames.Selected(0) = True
    frmPlayList.lstfiles.Selected(0) = True
End If
frmMain.lblPlaylist.Caption = GetSetting("mp3 d00d", "properties", "Playlist")
    If frmMain.lblPlaylist.Caption = "Close" Then
        frmPlayList.Visible = False
    Else
       frmMain.chkPlayList.Value = 1
       frmMain.chkPlayList_Click
    End If
End Sub
