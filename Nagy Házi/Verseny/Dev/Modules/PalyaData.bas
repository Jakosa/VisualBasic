Attribute VB_Name = "PalyaData"
Option Explicit

' P�ly�k k�nyvt�r�nak neve.
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

' P�lya inform�ci�k
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
    ' P�ly�hoz tartoz� ide�lis k�rsz�m.
    KorokSzama As Byte
    SorrendTomb() As Sorrend
    ' Aut�k be�ll�t�sait t�rol� t�mb.
    Autok(1 To 4) As New Auto
    ' Versenyp�ly�n l�v� aut�k sz�m�t t�rolja.
    AutokSzama As Byte
    KocsiVonalTomb() As VonalKoordinatak
    KocsiVonalakSzama As Integer
End Type

Public PalyaInfo As PInfo
' Alap�rtelmez�sben ennyir�l indul el.
Public Const KezdoSzektorido = 100000
' 5 m-t jelent. Ez azt jelenti hogy egy elmozdul�ssal az aut� 10 m�tert tesz meg.
Public Const PalyaHosszanakLepteke = 5

Public Function Vizsgalat() As Boolean
    If PalyaInfo.KocsiVonalakSzama - 1 < 4 Then
        WarningWindow "Hi�nyos adatok!", "Nincsen elegend� kocsi vonal l�trehozva! Jelenleg: " & CStr(PalyaInfo.KocsiVonalakSzama - 1) & ". Minimum: 4.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.KocsiVonalakSzama - 1 > 4 Then
        WarningWindow "Hi�nyos adatok!", "T�l sok kocsi vonal lett l�trehozva! Jelenleg: " & CStr(PalyaInfo.KocsiVonalakSzama - 1) & ". Maximum: 4.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.PalyaVonalakSzama = 0 Then
        WarningWindow "Hi�nyos adatok!", "Nincsenek p�lya vonalak!", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.SzektorNevekSzama < 3 Then
        WarningWindow "Hi�nyos adatok!", "Nincsen elegend� szektor n�v l�trehozva! Jelenleg: " & CStr(PalyaInfo.SzektorNevekSzama) & ". Minimum: 3.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.SzektorNevekSzama > 3 Then
        WarningWindow "Hi�nyos adatok!", "T�l sok szektor n�v lett l�trehozva! Jelenleg: " & CStr(PalyaInfo.SzektorNevekSzama) & ". Maximum: 3.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.SzektorVonalakSzama < 3 Then
        WarningWindow "Hi�nyos adatok!", "Nincsen elegend� szektor vonal l�trehozva! Jelenleg: " & CStr(PalyaInfo.SzektorVonalakSzama) & ". Minimum: 3.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.SzektorVonalakSzama > 3 Then
        WarningWindow "Hi�nyos adatok!", "T�l sok szektor n�v lett l�trehozva! Jelenleg: " & CStr(PalyaInfo.SzektorVonalakSzama) & ". Maximum: 3.", True
        Vizsgalat = False
        Exit Function
    End If

    If PalyaInfo.StartCelVonalNev.Label Is Nothing Then
        WarningWindow "Hi�nyos adatok!", "Nincsen start/c�lvonal l�trehozva!", True
        Vizsgalat = False
        Exit Function
    End If

    Vizsgalat = True
End Function
