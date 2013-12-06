Attribute VB_Name = "PalyaData"
' Fejléc
' Készítette: Jakosa Csaba Árpád
' Fejléc vége

Option Explicit

' Pályák könyvtárának neve.
Public Const MapDir = "Maps"

' A sorrendet tárolja.
Public Type SorrendAdatok
    ' Az autó szine.
    Szin As String
    ' Az autó ideje.
    Ido As Date
End Type

' A szektorokban mért autók sorrendjét tárolja.
Public Type SzektorSorrend
    ' Tárolja hogy már van-e adat vagy még nincs.
    VanAdat As Boolean
    ' Tárolja az egyes autókhoz tartozó sorrend adatait.
    Autok(1 To 4) As SorrendAdatok
End Type

' Tárolja a sorrendet.
Public Type Sorrend
    ' Tárolja a három szerktorban az autok sorrendjét.
    Szektor(1 To 3) As SzektorSorrend
End Type

' Tárolja hogy lehet-e új játékot indítani vagy sem.
Public NewGameEnabled As Boolean

' Vonalak koordinátáit tárolja.
Public Type VonalKoordinatak
    ' X1 pont.
    X1 As Integer
    ' X2 pont.
    X2 As Integer
    ' Y1 pont.
    Y1 As Integer
    ' Y2 pont.
    Y2 As Integer
    ' Tárolja a létrehozott vonalat.
    Vonal As VB.Line
End Type

' A név koordinátáit tárolja.
Public Type NevKoordinatak
    ' Left távolsága.
    Left As Integer
    ' Top távolsága.
    Top As Integer
    ' Tárolja a létrehozott "Label"-t.
    Label As VB.Label
End Type

' Pálya információk
Public Type PInfo
    ' Tárolja a pályák nevét.
    PalyaNevek() As String
    ' Tárolja a pálya körvonalát.
    PalyaVonalTomb() As VonalKoordinatak
    ' Tárolja a szektorok vonalát.
    SzektorVonalTomb() As VonalKoordinatak
    ' Tárolja a szektor neveket.
    SzektorNevTomb() As NevKoordinatak
    ' Tárolja a start/célvonal nevét.
    StartCelVonalNev As NevKoordinatak
    ' Tárolja a pálya neveinek számát.
    PalyaNevekSzama As Integer
    ' Tárolja a pálya körvonalainak számát.
    PalyaVonalakSzama As Integer
    ' Tárolja a szektorok vonalainak számát.
    SzektorVonalakSzama As Integer
    ' Tárolja a szektor neveinek számát.
    SzektorNevekSzama As Integer
    ' Pályához tartozó ideális körszám.
    KorokSzama As Byte
    ' Tárolja a kocsik sorrendjét.
    SorrendTomb() As Sorrend
    ' Autók beállításait tároló tömb.
    Autok(1 To 4) As New Auto
    ' Versenypályán lévõ autók számát tárolja.
    AutokSzama As Byte
    ' Tárolja az egyes kocsik vonalait (külsejét).
    KocsiVonalTomb() As VonalKoordinatak
    ' Tárolja hány darab kocsinak van vonala (külseje).
    KocsiVonalakSzama As Integer
End Type

' Tárolja a pálya információit.
Public PalyaInfo As PInfo
' Alapértelmezésben ennyirõl indul el.
Public Const KezdoSzektorido = 100000
' Alapértelmezésben ennyirõl indul a pontok közötti távolság számítás.
Public Const KezdoTavolsagPontok = 1000000
' 5 m-t jelent. Ez azt jelenti hogy egy elmozdulással az autó 10 métert tesz meg.
Public Const PalyaHosszanakLepteke = 5

' Megvizsgálja hogy minden adat megfelelõ-e a program normális használatához.
Public Function Vizsgalat() As Boolean
    ' Akkor fut le ha nincs elég kocsi vonal létrehozva.
    If PalyaInfo.KocsiVonalakSzama - 1 < 4 Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Hiányos adatok!", "Nincsen elegendõ kocsi vonal létrehozva! Jelenleg: " & CStr(PalyaInfo.KocsiVonalakSzama - 1) & ". Minimum: 4.", True
        ' Hamis értéket add vissza ha hiba van a program indításában.
        Vizsgalat = False
        ' Kilépés a függvénybõl.
        Exit Function
    End If

    ' Akkor fut le ha túl sok kocsi vonal van létrehozva.
    If PalyaInfo.KocsiVonalakSzama - 1 > 4 Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Hiányos adatok!", "Túl sok kocsi vonal lett létrehozva! Jelenleg: " & CStr(PalyaInfo.KocsiVonalakSzama - 1) & ". Maximum: 4.", True
        ' Hamis értéket add vissza ha hiba van a program indításában.
        Vizsgalat = False
        ' Kilépés a függvénybõl.
        Exit Function
    End If

    ' Akkor fut le ha nincs a pályának körvonala.
    If PalyaInfo.PalyaVonalakSzama = 0 Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Hiányos adatok!", "Nincsenek pálya vonalak!", True
        ' Hamis értéket add vissza ha hiba van a program indításában.
        Vizsgalat = False
        ' Kilépés a függvénybõl.
        Exit Function
    End If

    ' Akkor fut le ha nincs elegendõ szektor név létrehozva.
    If PalyaInfo.SzektorNevekSzama < 3 Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Hiányos adatok!", "Nincsen elegendõ szektor név létrehozva! Jelenleg: " & CStr(PalyaInfo.SzektorNevekSzama) & ". Minimum: 3.", True
        ' Hamis értéket add vissza ha hiba van a program indításában.
        Vizsgalat = False
        ' Kilépés a függvénybõl.
        Exit Function
    End If

    ' Akkor fut le ha túl sok szektor név van létrehozva.
    If PalyaInfo.SzektorNevekSzama > 3 Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Hiányos adatok!", "Túl sok szektor név lett létrehozva! Jelenleg: " & CStr(PalyaInfo.SzektorNevekSzama) & ". Maximum: 3.", True
        ' Hamis értéket add vissza ha hiba van a program indításában.
        Vizsgalat = False
        ' Kilépés a függvénybõl.
        Exit Function
    End If

    ' Akkor fut le ha kevés szektor vonal van létrehozva.
    If PalyaInfo.SzektorVonalakSzama < 3 Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Hiányos adatok!", "Nincsen elegendõ szektor vonal létrehozva! Jelenleg: " & CStr(PalyaInfo.SzektorVonalakSzama) & ". Minimum: 3.", True
        ' Hamis értéket add vissza ha hiba van a program indításában.
        Vizsgalat = False
        ' Kilépés a függvénybõl.
        Exit Function
    End If

    ' Akkor fut le ha túl sok szektor vonal van létrehozva.
    If PalyaInfo.SzektorVonalakSzama > 3 Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Hiányos adatok!", "Túl sok szektor név lett létrehozva! Jelenleg: " & CStr(PalyaInfo.SzektorVonalakSzama) & ". Maximum: 3.", True
        ' Hamis értéket add vissza ha hiba van a program indításában.
        Vizsgalat = False
        ' Kilépés a függvénybõl.
        Exit Function
    End If

    ' Akkor fut le ha nincs start/cél vonal létrehozva.
    If PalyaInfo.StartCelVonalNev.Label Is Nothing Then
        ' Figyelmeztetés ablak megnyítása.
        WarningWindow "Hiányos adatok!", "Nincsen start/célvonal létrehozva!", True
        ' Hamis értéket add vissza ha hiba van a program indításában.
        Vizsgalat = False
        ' Kilépés a függvénybõl.
        Exit Function
    End If

    ' Igaz értékket add vissza ha ninc shiba a program indulásában.
    Vizsgalat = True
End Function
