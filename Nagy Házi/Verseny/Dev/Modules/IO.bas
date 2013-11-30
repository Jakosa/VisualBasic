Attribute VB_Name = "IO"
Option Explicit

' Megnézi hogy könyvtár-e vagy sem.
' A "DirName" a fájl nevét tárolja.
Public Function IsDirectory(ByVal DirName As String) As Boolean
    ' Megnézi hogy könyvtár-e vagy sem.
    If GetAttr(DirName) And vbDirectory Then
        ' Visszadja igaz értékkel hogy könyvtár.
        IsDirectory = True
    End If
End Function

' Megnézi hogy fájl-e vagy sem.
' A "FileName" a fájl nevét tárolja.
Public Function IsFile(ByVal FileName As String) As Boolean
    ' Visszadja igaz vagy hamis értékkel hogy fájl-e.
    IsFile = (GetAttr(FileName) And vbDirectory) = 0
End Function

' Megnézi létezik-e az adott fájl.
' A "FileName" a fájl nevét tárolja.
Public Function FileExists(ByVal FileName As String) As Boolean
    ' Ha hiba van akkor ugrik a hibarészhez.
    On Error GoTo ErrorHandler

    ' Visszadja igaz vagy hamis értékkel létezik-e a fájl.
    FileExists = (GetAttr(FileName) And vbDirectory) = 0

    ' Hiba esetén ide ugrik a program.
ErrorHandler:
    ' Hamis (False) értékkel lép ki.
End Function

' Megnézi létezik-e az adott könyvtár.
' A "DirName" a könyvtár nevét tárolja.
Public Function DirExists(ByVal DirName As String) As Boolean
    ' Visszadja igaz vagy hamis értékkel létezik-e a könyvtár.
    DirExists = (Dir(DirName, vbDirectory) = "")
End Function
