Attribute VB_Name = "IO"
Option Explicit

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
