Attribute VB_Name = "IO"
Option Explicit

' Megn�zi hogy k�nyvt�r-e vagy sem.
' A "DirName" a f�jl nev�t t�rolja.
Public Function IsDirectory(ByVal DirName As String) As Boolean
    ' Megn�zi hogy k�nyvt�r-e vagy sem.
    If GetAttr(DirName) And vbDirectory Then
        ' Visszadja igaz �rt�kkel hogy k�nyvt�r.
        IsDirectory = True
    End If
End Function

' Megn�zi hogy f�jl-e vagy sem.
' A "FileName" a f�jl nev�t t�rolja.
Public Function IsFile(ByVal FileName As String) As Boolean
    ' Visszadja igaz vagy hamis �rt�kkel hogy f�jl-e.
    IsFile = (GetAttr(FileName) And vbDirectory) = 0
End Function

' Megn�zi l�tezik-e az adott f�jl.
' A "FileName" a f�jl nev�t t�rolja.
Public Function FileExists(ByVal FileName As String) As Boolean
    ' Ha hiba van akkor ugrik a hibar�szhez.
    On Error GoTo ErrorHandler

    ' Visszadja igaz vagy hamis �rt�kkel l�tezik-e a f�jl.
    FileExists = (GetAttr(FileName) And vbDirectory) = 0

    ' Hiba eset�n ide ugrik a program.
ErrorHandler:
    ' Hamis (False) �rt�kkel l�p ki.
End Function

' Megn�zi l�tezik-e az adott k�nyvt�r.
' A "DirName" a k�nyvt�r nev�t t�rolja.
Public Function DirExists(ByVal DirName As String) As Boolean
    ' Visszadja igaz vagy hamis �rt�kkel l�tezik-e a k�nyvt�r.
    DirExists = (Dir(DirName, vbDirectory) = "")
End Function
