Attribute VB_Name = "PalyaData"
Option Explicit

' Pályák könyvtárának neve.
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

' Pálya információk
Public Type PInfo
    PalyaNevek() As String
    PalyaVonalTomb() As VonalKoordinatak
    SzektorVonalTomb() As VonalKoordinatak
    SzektorNevTomb() As NevKoordinatak
    StartCelVonalNev As NevKoordinatak
    PalyaNevekSzama As Integer
    PalyaVonalakSzama As Integer
    SzektorVonalakSzama As Integer
    SzektorNevekSzama As Integer
    ' Pályához tartozó ideális körszám.
    KorokSzama As Byte
    SorrendTomb() As Sorrend
    ' Autók beállításait tároló tömb.
    Autok(1 To 4) As New Auto
    ' Versenypályán lévõ autók számát tárolja.
    AutokSzama As Byte
    KocsiVonalTomb() As VonalKoordinatak
    KocsiVonalakSzama As Integer
End Type

Public PalyaInfo As PInfo
' Alapértelmezésben ennyirõl indul el.
Public Const KezdoSzektorido = 100000
' 5 m-t jelent. Ez azt jelenti hogy egy elmozdulással az autó 10 métert tesz meg.
Public Const PalyaHosszanakLepteke = 5

Public Function Vizsgalat() As Boolean
    If PalyaInfo.KocsiVonalakSzama - 1 < 4 Then
        WarningWindow "Hiányos adatok!", "Nincsen elegendõ kocsi vonal létrehozva! Jelenleg: " & CStr(PalyaInfo.KocsiVonalakSzama - 1) & ". Minimum: 4.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.KocsiVonalakSzama - 1 > 4 Then
        WarningWindow "Hiányos adatok!", "Túl sok kocsi vonal lett létrehozva! Jelenleg: " & CStr(PalyaInfo.KocsiVonalakSzama - 1) & ". Maximum: 4.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.PalyaVonalakSzama = 0 Then
        WarningWindow "Hiányos adatok!", "Nincsenek pálya vonalak!", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.SzektorNevekSzama < 3 Then
        WarningWindow "Hiányos adatok!", "Nincsen elegendõ szektor név létrehozva! Jelenleg: " & CStr(PalyaInfo.SzektorNevekSzama) & ". Minimum: 3.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.SzektorNevekSzama > 3 Then
        WarningWindow "Hiányos adatok!", "Túl sok szektor név lett létrehozva! Jelenleg: " & CStr(PalyaInfo.SzektorNevekSzama) & ". Maximum: 3.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.SzektorVonalakSzama < 3 Then
        WarningWindow "Hiányos adatok!", "Nincsen elegendõ szektor vonal létrehozva! Jelenleg: " & CStr(PalyaInfo.SzektorVonalakSzama) & ". Minimum: 3.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.SzektorVonalakSzama > 3 Then
        WarningWindow "Hiányos adatok!", "Túl sok szektor név lett létrehozva! Jelenleg: " & CStr(PalyaInfo.SzektorVonalakSzama) & ". Maximum: 3.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.StartCelVonalNev.Label Is Nothing Then
        WarningWindow "Hiányos adatok!", "Nincsen start/célvonal létrehozva!", True
        Vizsgalat = False
        Exit Function
    End If

    Vizsgalat = True
End Function
