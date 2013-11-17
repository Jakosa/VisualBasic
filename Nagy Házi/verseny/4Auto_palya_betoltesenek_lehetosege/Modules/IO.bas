Attribute VB_Name = "IO"
Option Explicit

Public Function IsDirectory(ByVal Name As String) As Boolean
    If GetAttr(Name) And vbDirectory Then
        IsDirectory = True
    End If
End Function

Public Function IsFile(ByVal Name As String) As Boolean
    IsFile = (GetAttr(Name) And vbDirectory) = 0
End Function

Public Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler

    ' Megn�zi l�tezik-e vagy sem.
    FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
    ' False �rt�kkel l�p ki
End Function

Public Function DirExists(DirName As String) As Boolean
    DirExists = (Dir(DirName, vbDirectory) = "")
End Function
