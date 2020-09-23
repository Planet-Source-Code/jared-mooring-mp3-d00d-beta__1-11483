Attribute VB_Name = "basFiles"
Public Sub OpenPathList(path As String, Form As Form)
'===============================
'used for retrieving the contents of the files saved
'for the users playlist
'===============================

Filehandle = FreeFile
Dim strJared As String
Open path For Input As #Filehandle
On Error GoTo shit2
shit:
       Input #Filehandle, strJared
       
            Form.lstfiles.AddItem strJared
            Form.frmMain.lstfiles.AddItem strJared
       
      GoTo shit
shit2:
Close #Filehandle
End Sub
Public Sub SaveFileList(path As String, Form As Form)
'======================================
'used for saving the names of the files in the playlist
'======================================
On Error Resume Next
Dim Save As Long
Dim fFile As Integer
fFile = FreeFile
    Open path For Output As #1
        For Save = 0 To Form.lstNames.ListCount - 1
            Print #1, Form.lstNames.List(Save)
        Next Save
    Close #1
End Sub

Public Sub OpenFileList(path As String, Form As Form)
'=========================================
'used for opening the contents of the file list
'and putting the contents onto the main playlist
'=========================================
On Error Resume Next

Filehandle = FreeFile
Dim strJared As String
Open path For Input As #Filehandle
On Error GoTo shit2
shit:
       Input #Filehandle, strJared
            Form.lstNames.AddItem strJared
            Form.frmMain.lstNames.AddItem strJared
            GoTo shit
shit2:
        Close #Filehandle
End Sub
Public Sub SavePathList(path As String, Form As Form)
'======================================
'used for saing the path names into the users playlist
'======================================

On Error Resume Next
Dim Save As Long
Dim fFile As Integer
fFile = FreeFile
    Open path For Output As #1
        For Save = 0 To Form.lstfiles.ListCount - 1
            Print #1, Form.lstfiles.List(Save)
        Next Save
    Close #1
End Sub

