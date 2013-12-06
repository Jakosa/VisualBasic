Attribute VB_Name = "PalyaData"
' Fejl�c
' K�sz�tette: Jakosa Csaba �rp�d
' Fejl�c v�ge

Option Explicit

' P�ly�k k�nyvt�r�nak neve.
Public Const MapDir = "Maps"

' A sorrendet t�rolja.
Public Type SorrendAdatok
    ' Az aut� szine.
    Szin As String
    ' Az aut� ideje.
    Ido As Date
End Type

' A szektorokban m�rt aut�k sorrendj�t t�rolja.
Public Type SzektorSorrend
    ' T�rolja hogy m�r van-e adat vagy m�g nincs.
    VanAdat As Boolean
    ' T�rolja az egyes aut�khoz tartoz� sorrend adatait.
    Autok(1 To 4) As SorrendAdatok
End Type

' T�rolja a sorrendet.
Public Type Sorrend
    ' T�rolja a h�rom szerktorban az autok sorrendj�t.
    Szektor(1 To 3) As SzektorSorrend
End Type

' T�rolja hogy lehet-e �j j�t�kot ind�tani vagy sem.
Public NewGameEnabled As Boolean

' Vonalak koordin�t�it t�rolja.
Public Type VonalKoordinatak
    ' X1 pont.
    X1 As Integer
    ' X2 pont.
    X2 As Integer
    ' Y1 pont.
    Y1 As Integer
    ' Y2 pont.
    Y2 As Integer
    ' T�rolja a l�trehozott vonalat.
    Vonal As VB.Line
End Type

' A n�v koordin�t�it t�rolja.
Public Type NevKoordinatak
    ' Left t�vols�ga.
    Left As Integer
    ' Top t�vols�ga.
    Top As Integer
    ' T�rolja a l�trehozott "Label"-t.
    Label As VB.Label
End Type

' P�lya inform�ci�k
Public Type PInfo
    ' T�rolja a p�ly�k nev�t.
    PalyaNevek() As String
    ' T�rolja a p�lya k�rvonal�t.
    PalyaVonalTomb() As VonalKoordinatak
    ' T�rolja a szektorok vonal�t.
    SzektorVonalTomb() As VonalKoordinatak
    ' T�rolja a szektor neveket.
    SzektorNevTomb() As NevKoordinatak
    ' T�rolja a start/c�lvonal nev�t.
    StartCelVonalNev As NevKoordinatak
    ' T�rolja a p�lya neveinek sz�m�t.
    PalyaNevekSzama As Integer
    ' T�rolja a p�lya k�rvonalainak sz�m�t.
    PalyaVonalakSzama As Integer
    ' T�rolja a szektorok vonalainak sz�m�t.
    SzektorVonalakSzama As Integer
    ' T�rolja a szektor neveinek sz�m�t.
    SzektorNevekSzama As Integer
    ' P�ly�hoz tartoz� ide�lis k�rsz�m.
    KorokSzama As Byte
    ' T�rolja a kocsik sorrendj�t.
    SorrendTomb() As Sorrend
    ' Aut�k be�ll�t�sait t�rol� t�mb.
    Autok(1 To 4) As New Auto
    ' Versenyp�ly�n l�v� aut�k sz�m�t t�rolja.
    AutokSzama As Byte
    ' T�rolja az egyes kocsik vonalait (k�lsej�t).
    KocsiVonalTomb() As VonalKoordinatak
    ' T�rolja h�ny darab kocsinak van vonala (k�lseje).
    KocsiVonalakSzama As Integer
End Type

' T�rolja a p�lya inform�ci�it.
Public PalyaInfo As PInfo
' Alap�rtelmez�sben ennyir�l indul el.
Public Const KezdoSzektorido = 100000
' Alap�rtelmez�sben ennyir�l indul a pontok k�z�tti t�vols�g sz�m�t�s.
Public Const KezdoTavolsagPontok = 1000000
' 5 m-t jelent. Ez azt jelenti hogy egy elmozdul�ssal az aut� 10 m�tert tesz meg.
Public Const PalyaHosszanakLepteke = 5

' Megvizsg�lja hogy minden adat megfelel�-e a program norm�lis haszn�lat�hoz.
Public Function Vizsgalat() As Boolean
    ' Akkor fut le ha nincs el�g kocsi vonal l�trehozva.
    If PalyaInfo.KocsiVonalakSzama - 1 < 4 Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Hi�nyos adatok!", "Nincsen elegend� kocsi vonal l�trehozva! Jelenleg: " & CStr(PalyaInfo.KocsiVonalakSzama - 1) & ". Minimum: 4.", True
        ' Hamis �rt�ket add vissza ha hiba van a program ind�t�s�ban.
        Vizsgalat = False
        ' Kil�p�s a f�ggv�nyb�l.
        Exit Function
    End If

    ' Akkor fut le ha t�l sok kocsi vonal van l�trehozva.
    If PalyaInfo.KocsiVonalakSzama - 1 > 4 Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Hi�nyos adatok!", "T�l sok kocsi vonal lett l�trehozva! Jelenleg: " & CStr(PalyaInfo.KocsiVonalakSzama - 1) & ". Maximum: 4.", True
        ' Hamis �rt�ket add vissza ha hiba van a program ind�t�s�ban.
        Vizsgalat = False
        ' Kil�p�s a f�ggv�nyb�l.
        Exit Function
    End If

    ' Akkor fut le ha nincs a p�ly�nak k�rvonala.
    If PalyaInfo.PalyaVonalakSzama = 0 Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Hi�nyos adatok!", "Nincsenek p�lya vonalak!", True
        ' Hamis �rt�ket add vissza ha hiba van a program ind�t�s�ban.
        Vizsgalat = False
        ' Kil�p�s a f�ggv�nyb�l.
        Exit Function
    End If

    ' Akkor fut le ha nincs elegend� szektor n�v l�trehozva.
    If PalyaInfo.SzektorNevekSzama < 3 Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Hi�nyos adatok!", "Nincsen elegend� szektor n�v l�trehozva! Jelenleg: " & CStr(PalyaInfo.SzektorNevekSzama) & ". Minimum: 3.", True
        ' Hamis �rt�ket add vissza ha hiba van a program ind�t�s�ban.
        Vizsgalat = False
        ' Kil�p�s a f�ggv�nyb�l.
        Exit Function
    End If

    ' Akkor fut le ha t�l sok szektor n�v van l�trehozva.
    If PalyaInfo.SzektorNevekSzama > 3 Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Hi�nyos adatok!", "T�l sok szektor n�v lett l�trehozva! Jelenleg: " & CStr(PalyaInfo.SzektorNevekSzama) & ". Maximum: 3.", True
        ' Hamis �rt�ket add vissza ha hiba van a program ind�t�s�ban.
        Vizsgalat = False
        ' Kil�p�s a f�ggv�nyb�l.
        Exit Function
    End If

    ' Akkor fut le ha kev�s szektor vonal van l�trehozva.
    If PalyaInfo.SzektorVonalakSzama < 3 Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Hi�nyos adatok!", "Nincsen elegend� szektor vonal l�trehozva! Jelenleg: " & CStr(PalyaInfo.SzektorVonalakSzama) & ". Minimum: 3.", True
        ' Hamis �rt�ket add vissza ha hiba van a program ind�t�s�ban.
        Vizsgalat = False
        ' Kil�p�s a f�ggv�nyb�l.
        Exit Function
    End If

    ' Akkor fut le ha t�l sok szektor vonal van l�trehozva.
    If PalyaInfo.SzektorVonalakSzama > 3 Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Hi�nyos adatok!", "T�l sok szektor n�v lett l�trehozva! Jelenleg: " & CStr(PalyaInfo.SzektorVonalakSzama) & ". Maximum: 3.", True
        ' Hamis �rt�ket add vissza ha hiba van a program ind�t�s�ban.
        Vizsgalat = False
        ' Kil�p�s a f�ggv�nyb�l.
        Exit Function
    End If

    ' Akkor fut le ha nincs start/c�l vonal l�trehozva.
    If PalyaInfo.StartCelVonalNev.Label Is Nothing Then
        ' Figyelmeztet�s ablak megny�t�sa.
        WarningWindow "Hi�nyos adatok!", "Nincsen start/c�lvonal l�trehozva!", True
        ' Hamis �rt�ket add vissza ha hiba van a program ind�t�s�ban.
        Vizsgalat = False
        ' Kil�p�s a f�ggv�nyb�l.
        Exit Function
    End If

    ' Igaz �rt�kket add vissza ha ninc shiba a program indul�s�ban.
    Vizsgalat = True
End Function
