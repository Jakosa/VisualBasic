Attribute VB_Name = "PalyaData"
Option Explicit

Public Const MapDir = "Maps"

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

Public Type VonalKoordinatak
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    Vonal As VB.Line
End Type

Public Type NevKoordinatak
    Left As Integer
    Top As Integer
    Label As VB.Label
End Type

Public Type PInfo ' Pálya információ
    PalyaNevek() As String
    PalyaVonalTomb() As VonalKoordinatak
    SzektorVonalTomb() As VonalKoordinatak
    SzektorNevTomb() As NevKoordinatak
    StartCelVonalNev As NevKoordinatak
    PalyaNevekSzama As Integer
    PalyaVonalakSzama As Integer
    SzektorVonalakSzama As Integer
    SzektorNevekSzama As Integer
    KorokSzama As Byte ' Pályához tartozó ideális körszám.
End Type

Public PalyaInfo As PInfo

