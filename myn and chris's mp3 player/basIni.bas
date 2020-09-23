Attribute VB_Name = "basIni"
'Source:    http://www.planetsourcecode.com
'Auhtor:    Unknown

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public ret As String

Public Sub WriteINI(FileName As String, Section As String, Key As String, Text As String)
WritePrivateProfileString Section, Key, Text, FileName
End Sub

Public Function ReadINI(FileName As String, Section As String, Key As String)
ret = Space$(255)
RetLen = GetPrivateProfileString(Section, Key, "", ret, Len(ret), FileName)
ret = Left$(ret, RetLen)
ReadINI = ret
End Function
Public Function GetFileTitle(ByVal sFilename As String) As String
    'Returns FileTitle Without Path
    Dim lPos As Long
    lPos = InStrRev(sFilename, "\")


    If lPos > 0 Then


        If lPos < Len(sFilename) Then
            GetFileTitle = Mid$(sFilename, lPos + 1)
        Else
            GetFileTitle = ""
        End If
    Else
        GetFileTitle = sFilename
    End If
    
End Function


