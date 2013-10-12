Attribute VB_Name = "PalyaData"
Option Explicit

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
