Attribute VB_Name = "PalyaData"
Option Explicit

Public Type SzektorSorrend
    VanAdat As Boolean
    AutoSzine(1 To 4) As String
End Type

Public Type Sorrend
    Dist As Single
    Szin As String
    Szektor(1 To 3) As SzektorSorrend
End Type

Public SorrendTomb() As Sorrend
Public Autok(1 To 4) As New Auto   ' Autók beállítását tároló tömb.
