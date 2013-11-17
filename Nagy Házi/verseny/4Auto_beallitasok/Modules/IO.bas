Attribute VB_Name = "IO"
Option Explicit

Public Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler

    ' Megnézi létezik-e vagy sem.
    FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
    ' False értékkel lép ki
End Function

Public Function DirExists(DirName As String) As Boolean
    DirExists = (Dir(DirName, vbDirectory) = "")
End Function
