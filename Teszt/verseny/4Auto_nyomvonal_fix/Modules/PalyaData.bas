Attribute VB_Name = "PalyaData"
Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function SetViewportOrgEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Public Const WM_PAINT As Long = &HF&

Public Type SorrendAdatok
    Szin As String
    Ido As Date
End Type

Public Type SzektorSorrend
    VanAdat As Boolean
    Autok(1 To 4) As SorrendAdatok
End Type

Public Type Sorrend
    Szektor(1 To 3) As SzektorSorrend
End Type

Public SorrendTomb() As Sorrend
Public Autok(1 To 4) As New Auto   ' Autók beállítását tároló tömb.

Public NewGameEnabled As Boolean
