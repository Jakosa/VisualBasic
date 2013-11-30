VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Pálya elérési útvonalát tárolja.
Private MapPath As String

' Beállítjuk a form létrehozásakor az alap folyamatokat.
Private Sub Class_Initialize()
    ' Beállítja az elérési útvonalat.
    MapPath = MapDir & "/" & Config.Globalis_PalyaNeve
End Sub

' Pálya betöltése.
' A "MapName" tárolja a pálya fájlának nevét.
Public Sub LoadMap(ByVal MapName As String)
    ' Kitakarítja a változókat.
    Clean

    ' Megnézzi létezik-e a pálya mappája. Ha nem akkor létrehozza.
    If DirExists(MapDir) Then
        ' Pálya mappa létrehozása.
        MkDir MapDir
    End If

    ' Beállítja az elérési útvonalat.
    MapPath = MapDir & "/" & MapName

    ' Megnézzi létezik-e a pálya fájl.
    If FileExists(MapPath) Then
    ' Ha új pálya fájl lenne létrehozva akkor is elfut ez a rész hogy betöltödjenek a beállítások.
UjBetoltese:
        ' Fájl megnyítása olvasásra.
        Open MapPath For Input As #1

        ' Addig fog futni amíg nem ér a konfig fájl végére.
        Do Until EOF(1)
            ' A program soronként fogja az adatokat betölteni így a változó a sorok adatát fogja tárolni.
            Dim Sor As String
            ' Adatok betöltése soronként.
            Line Input #1, Sor

            ' Csak akkor fut le ha van valami szöveg az adott sorban.
            If Len(Trim(Sor)) > 0 Then
                ' Kiszüri a kommenteket ezzel a feltételel. Így csak akkor fut le ha nem kommentes sor.
                If Not Mid$(Sor, 1, 1) = "#" Then
                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "PalyaVonal") > 0 Then
                        ' Ha nagyobb a pálya vonalak száma mint amit a tömb tárolni tud akkor fut le.
                        If PalyaInfo.PalyaVonalakSzama >= UBound(PalyaInfo.PalyaVonalTomb) Then
                            ' Tömb megnövelése.
                            ReDim Preserve PalyaInfo.PalyaVonalTomb(0 To UBound(PalyaInfo.PalyaVonalTomb) + 50) As VonalKoordinatak
                        End If

                        ' Beállítás betöltése.
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama) = ParseLine(Parse(Sor))
                        ' Vonal létrehozása.
                        Set PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal = Palya.Controls.Add("VB.Line", "PalyaVonal" & CStr(PalyaInfo.PalyaVonalakSzama), Palya)
                        'PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.BorderColor = &H80000008
                        ' X1 pont beállítása.
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.X1 = PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).X1
                        ' X2 pont beállítása.
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.X2 = PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).X2
                        ' Y1 pont beállítása.
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.Y1 = PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Y1
                        ' Y2 pont beállítása.
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.Y2 = PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Y2
                        ' Vonal szélességének beállítása.
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.BorderWidth = 2
                        ' Vonal láthatóvá tétele.
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.Visible = True
                        ' Érték megnövelése eggyel.
                        PalyaInfo.PalyaVonalakSzama = PalyaInfo.PalyaVonalakSzama + 1
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "SzektorVonal") > 0 Then
                        ' Ha nagyobb a szektor vonalak száma mint amit a tömb tárolni tud akkor fut le.
                        If PalyaInfo.SzektorVonalakSzama >= UBound(PalyaInfo.SzektorVonalTomb) Then
                            ' Tömb megnövelése.
                            ReDim Preserve PalyaInfo.SzektorVonalTomb(0 To UBound(PalyaInfo.SzektorVonalTomb) + 5) As VonalKoordinatak
                        End If

                        ' Beállítás betöltése.
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama) = ParseLine(Parse(Sor))
                        ' Vonal létrehozása.
                        Set PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal = Palya.Controls.Add("VB.Line", "SzektorVonal" & CStr(PalyaInfo.SzektorVonalakSzama), Palya)
                        'PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.BorderColor = &H80000008
                        ' X1 pont beállítása.
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.X1 = PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).X1
                        ' X2 pont beállítása.
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.X2 = PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).X2
                        ' Y1 pont beállítása.
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.Y1 = PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Y1
                        ' Y2 pont beállítása.
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.Y2 = PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Y2
                        ' Vonal szélességének beállítása.
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.BorderWidth = 1
                        ' Vonal láthatóvá tétele.
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.Visible = True
                        ' Érték megnövelése eggyel.
                        PalyaInfo.SzektorVonalakSzama = PalyaInfo.SzektorVonalakSzama + 1
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "SzektorNev") > 0 Then
                        ' Ha nagyobb a szektor neveinek száma mint amit a tömb tárolni tud akkor fut le.
                        If PalyaInfo.SzektorNevekSzama >= UBound(PalyaInfo.SzektorNevTomb) Then
                            ' Tömb megnövelése.
                            ReDim Preserve PalyaInfo.SzektorNevTomb(0 To UBound(PalyaInfo.SzektorNevTomb) + 5) As NevKoordinatak
                        End If

                        ' Beállítás betöltése.
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama) = ParseLabel(Parse(Sor))
                        ' "Label" létrehozása.
                        Set PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label = Palya.Controls.Add("VB.Label", "SzektorNev" & CStr(PalyaInfo.SzektorNevekSzama), Palya)
                        ' "Label" háttérszinének beállítása a pályához.
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.BackColor = Palya.BackColor
                        ' "Label" méretének automatikus változtatásának bekapcsolása.
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.AutoSize = True
                        ' "Label" szövegének beállítása.
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.Caption = "Szektor " & CStr(PalyaInfo.SzektorNevekSzama + 1)
                        ' "Left" pozició beállítása.
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.Left = PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Left
                        ' "Top" pozició beállítása.
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.Top = PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Top
                        ' "Label" láthatóvá tétele.
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.Visible = True
                        ' Érték megnövelése eggyel.
                        PalyaInfo.SzektorNevekSzama = PalyaInfo.SzektorNevekSzama + 1
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "StartCelVonalNev") > 0 Then
                        ' Beállítás betöltése.
                        PalyaInfo.StartCelVonalNev = ParseLabel(Parse(Sor))
                        Set PalyaInfo.StartCelVonalNev.Label = Palya.Controls.Add("VB.Label", "StartCelVonalNev", Palya)
                        ' "Label" háttérszinének beállítása a pályához.
                        PalyaInfo.StartCelVonalNev.Label.BackColor = Palya.BackColor
                        ' "Label" méretének automatikus változtatásának bekapcsolása.
                        PalyaInfo.StartCelVonalNev.Label.AutoSize = True
                        ' "Label" szövegének beállítása.
                        PalyaInfo.StartCelVonalNev.Label.Caption = "Start / Cél"
                        ' "Left" pozició beállítása.
                        PalyaInfo.StartCelVonalNev.Label.Left = PalyaInfo.StartCelVonalNev.Left
                        ' "Top" pozició beállítása.
                        PalyaInfo.StartCelVonalNev.Label.Top = PalyaInfo.StartCelVonalNev.Top
                        ' "Label" láthatóvá tétele.
                        PalyaInfo.StartCelVonalNev.Label.Visible = True
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "KorokSzama") > 0 Then
                        ' Megvizsgálja hogy "byte"-e az adat.
                        If IsByte(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            PalyaInfo.KorokSzama = CByte(Parse(Sor))
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            PalyaInfo.KorokSzama = AGlobalis_KorokSzama
                        End If

                        ' Ha nagyobb a körök száma mint a maximális vagy a minimálisnál kisebb akkor az alapértelmezetett tölti be.
                        If PalyaInfo.KorokSzama > 20 Or PalyaInfo.KorokSzama < 2 Then
                            ' Alapértelmezett beállítás betöltése.
                            PalyaInfo.KorokSzama = AGlobalis_KorokSzama
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "KocsiVonal") > 0 Then
                        ' Ha nagyobb a kocsi vonalak száma mint amit a tömb tárolni tud akkor fut le.
                        If PalyaInfo.KocsiVonalakSzama >= UBound(PalyaInfo.SzektorVonalTomb) Then
                            ' Tömb megnövelése.
                            ReDim Preserve PalyaInfo.KocsiVonalTomb(1 To UBound(PalyaInfo.KocsiVonalTomb) + 4) As VonalKoordinatak
                        End If

                        ' Beállítás betöltése.
                        PalyaInfo.KocsiVonalTomb(PalyaInfo.KocsiVonalakSzama) = ParseLine(Parse(Sor))
                        ' Érték megnövelése eggyel.
                        PalyaInfo.KocsiVonalakSzama = PalyaInfo.KocsiVonalakSzama + 1
                    End If
                End If
            End If
        Loop

        ' Fájl lezására.
        Close #1
    Else
        ' Beállítja az elérési útvonalat.
        MapPath = MapDir & "/" & AGlobalis_PalyaNeve

        ' Ha nem létezik az alapértelmezett pálya akkor létrehozza.
        If Not FileExists(MapPath) Then
            ' Pálya létrehozása.
            CreateMap
        End If

        ' Pálya beállítások betöltése.
        GoTo UjBetoltese
    End If
End Sub

' Alapértelmezett pálya törlése.
Public Sub DeleteDefaultMap()
    ' Beállítja az elérési útvonalat.
    MapPath = MapDir & "/" & AGlobalis_PalyaNeve

    ' Megnézi létezik-e az alapértelmezett pálya. Ha igen akkor törli.
    If FileExists(MapPath) Then
        ' Pálya fájl törlése.
        Kill MapPath
    End If
End Sub

' Pálya fájl létrehozása.
Private Sub CreateMap()
    ' Fájl megnyítása írásra.
    Open MapPath For Output As #1

    ' Adatok kiírása a fájlba.
    Print #1, "# Alapértelmezett map."
    Print #1, ""
    Print #1, "# Pályához tartozó körök száma. Maximum: 20, Minimum: 2"
    Print #1, "KorokSzama: 5"
    Print #1, ""
    Print #1, "# Start és célvonal neve."
    Print #1, "StartCelVonalNev: 1680;4440"
    Print #1, ""
    Print #1, "# Szektorok neve."
    Print #1, "SzektorNev: 4680;360"
    Print #1, "SzektorNev: 6600;5640"
    Print #1, "SzektorNev: 0;4440"
    Print #1, ""
    Print #1, "# Szektor vonalak."
    Print #1, "SzektorVonal: 4560;3600;360;2280"
    Print #1, "SzektorVonal: 6480;5040;5880;4920"
    Print #1, "SzektorVonal: 1920;600;4680;4680"
    Print #1, ""
    Print #1, "# Kocsi vonalak."
    Print #1, "KocsiVonal: 1100;0;5000;0"
    Print #1, "KocsiVonal: 900;0;5000;0"
    Print #1, "KocsiVonal: 900;0;5300;0"
    Print #1, "KocsiVonal: 1100;0;5300;0"
    Print #1, ""
    Print #1, "# Pálya vonalak."
    Print #1, "PalyaVonal: 720;720;3720;4320"
    Print #1, "PalyaVonal: 960;720;3120;3720"
    Print #1, "PalyaVonal: 5040;4200;5640;6240"
    Print #1, "PalyaVonal: 4200;3360;6240;6480"
    Print #1, "PalyaVonal: 2880;2040;7080;6720"
    Print #1, "PalyaVonal: 2040;720;6720;5520"
    Print #1, "PalyaVonal: 5400;4680;6720;6960"
    Print #1, "PalyaVonal: 5760;5400;6000;6720"
    Print #1, "PalyaVonal: 4680;3720;6960;7200"
    Print #1, "PalyaVonal: 2280;3360;6000;6480"
    Print #1, "PalyaVonal: 1560;2280;5400;6000"
    Print #1, "PalyaVonal: 3720;2880;7200;7080"
    Print #1, "PalyaVonal: 5760;5400;4200;4920"
    Print #1, "PalyaVonal: 5400;5040;4920;5640"
    Print #1, "PalyaVonal: 6960;5760;4320;4200"
    Print #1, "PalyaVonal: 7200;6120;4920;5280"
    Print #1, "PalyaVonal: 6120;5760;5280;6000"
    Print #1, "PalyaVonal: 7680;6960;3720;4320"
    Print #1, "PalyaVonal: 8040;7200;4560;4920"
    Print #1, "PalyaVonal: 8400;8040;3840;4560"
    Print #1, "PalyaVonal: 8400;8400;3240;3840"
    Print #1, "PalyaVonal: 7680;7680;3120;3720"
    Print #1, "PalyaVonal: 7680;4560;3120;1920"
    Print #1, "PalyaVonal: 8400;4200;2640;840"
    Print #1, "PalyaVonal: 4560;3720;1920;1560"
    Print #1, "PalyaVonal: 4200;3360;840;960"
    Print #1, "PalyaVonal: 3720;3360;1560;1800"
    Print #1, "PalyaVonal: 3360;3240;1800;1920"
    Print #1, "PalyaVonal: 3240;3000;1920;2520"
    Print #1, "PalyaVonal: 2400;1800;3120;3240"
    Print #1, "PalyaVonal: 1800;1560;3240;3600"
    Print #1, "PalyaVonal: 1560;1560;4800;5400"
    Print #1, "PalyaVonal: 1560;1560;3600;4200"
    Print #1, "PalyaVonal: 3000;2400;2520;3120"
    Print #1, "PalyaVonal: 2040;1440;2520;2640"
    Print #1, "PalyaVonal: 2640;2400;1680;2280"
    Print #1, "PalyaVonal: 2400;2040;2280;2520"
    Print #1, "PalyaVonal: 3360;2880;960;1200"
    Print #1, "PalyaVonal: 8400;8400;2640;3240"
    Print #1, "PalyaVonal: 1440;960;2640;3120"
    Print #1, "PalyaVonal: 1560;1560;4200;4800"
    Print #1, "PalyaVonal: 2880;2640;1200;1680"
    Print #1, "PalyaVonal: 720;720;4320;4920"
    Print #1, "PalyaVonal: 720;720;4920;5520"

    ' Fájl lezására.
    Close #1
End Sub

' Szétvágja a betöltött sorokat. Majd visszadja az értékes adatokat.
' A "Sor" változó a pálya fájl egy sorát tartalmazza ami éppen kiolvasásra került.
Private Function Parse(ByVal Sor As String) As String
    ' Kettõspont jel utáni részt adja vissza. A felesleges adatokat eltávolítja (pl: szóköz).
    Parse = LTrim(Mid$(Sor, InStr(1, Sor, ":") + 1))
End Function

' Szétvágja a betöltött sorokat. Majd visszadja az értékes adatokat.
' A "Sor" változó a pálya fájl egy sorát tartalmazza ami éppen kiolvasásra került.
Private Function ParseLine(ByVal Sor As String) As VonalKoordinatak
    ' Változó létrehozása.
    Dim vk As VonalKoordinatak
    ' Üres tömb létrehozása.
    Dim v() As String

    ' Darabolja a sort kettõspont alapján.
    v = Split(Sor, ";")

    ' Akkor fut le ha nincs annyi eleme a tömbnek mint amennyi a hibátlan lefutáshoz kell.
    If UBound(v) - LBound(v) + 1 < 4 Then
        ' Kilépés a függvénybõl.
        Exit Function
    End If

    ' Megnézi hogy a változó "integer"-e.
    If IsInteger(v(0)) Then
        ' Érték átadása.
        vk.X1 = v(0)
    Else
        ' Nullára állítás.
        vk.X1 = 0
    End If

    ' Megnézi hogy a változó "integer"-e.
    If IsInteger(v(1)) Then
        ' Érték átadása.
        vk.X2 = v(1)
    Else
        ' Nullára állítás.
        vk.X2 = 0
    End If

    ' Megnézi hogy a változó "integer"-e.
    If IsInteger(v(2)) Then
        ' Érték átadása.
        vk.Y1 = v(2)
    Else
        ' Nullára állítás.
        vk.Y1 = 0
    End If

    ' Megnézi hogy a változó "integer"-e.
    If IsInteger(v(3)) Then
        ' Érték átadása.
        vk.Y2 = v(3)
    Else
        ' Nullára állítás.
        vk.Y2 = 0
    End If

    ' Beállított adatok visszadása.
    ParseLine = vk
End Function

' Szétvágja a betöltött sorokat. Majd visszadja az értékes adatokat.
' A "Sor" változó a pálya fájl egy sorát tartalmazza ami éppen kiolvasásra került.
Private Function ParseLabel(ByVal Sor As String) As NevKoordinatak
    ' Változó létrehozása.
    Dim nk As NevKoordinatak
    ' Üres tömb létrehozása.
    Dim v() As String

    ' Darabolja a sort kettõspont alapján.
    v = Split(Sor, ";")

    ' Akkor fut le ha nincs annyi eleme a tömbnek mint amennyi a hibátlan lefutáshoz kell.
    If UBound(v) - LBound(v) + 1 < 2 Then
        ' Kilépés a függvénybõl.
        Exit Function
    End If

    ' Megnézi hogy a változó "integer"-e.
    If IsInteger(v(0)) Then
        ' Érték átadása.
        nk.Left = v(0)
    Else
        ' Nullára állítás.
        nk.Left = 0
    End If

    ' Megnézi hogy a változó "integer"-e.
    If IsInteger(v(1)) Then
        ' Érték átadása.
        nk.Top = v(1)
    Else
        ' Nullára állítás.
        nk.Top = 0
    End If

    ' Beállított adatok visszadása.
    ParseLabel = nk
End Function

' Kitakarítja a változókat.
Private Sub Clean()
    ' "i" segédváltozó a ciklushoz.
    Dim i As Integer

    ' Ha nagyobb mint nulla akkor fut le.
    If PalyaInfo.PalyaVonalakSzama > 0 Then
        For i = LBound(PalyaInfo.PalyaVonalTomb) To PalyaInfo.PalyaVonalakSzama - 1
            ' Pályáról eltávolítja a vonalat.
            Palya.Controls.Remove PalyaInfo.PalyaVonalTomb(i).Vonal
            ' Nullázás.
            Set PalyaInfo.PalyaVonalTomb(i).Vonal = Nothing
        Next i
    End If

    ' Ha nagyobb mint nulla akkor fut le.
    If PalyaInfo.SzektorVonalakSzama > 0 Then
        For i = LBound(PalyaInfo.SzektorVonalTomb) To PalyaInfo.SzektorVonalakSzama - 1
            ' Pályáról eltávolítja a vonalat.
            Palya.Controls.Remove PalyaInfo.SzektorVonalTomb(i).Vonal
            ' Nullázás.
            Set PalyaInfo.SzektorVonalTomb(i).Vonal = Nothing
        Next i
    End If

    ' Ha nagyobb mint nulla akkor fut le.
    If PalyaInfo.SzektorNevekSzama > 0 Then
        For i = LBound(PalyaInfo.SzektorNevTomb) To PalyaInfo.SzektorNevekSzama - 1
            ' Pályáról eltávolítja a "Label"-t.
            Palya.Controls.Remove PalyaInfo.SzektorNevTomb(i).Label
            ' Nullázás.
            Set PalyaInfo.SzektorNevTomb(i).Label = Nothing
        Next i
    End If

    ' Ha nem nulla akkor fut le.
    If Not PalyaInfo.StartCelVonalNev.Label Is Nothing Then
        ' Pályáról eltávolítja a "Label"-t.
        Palya.Controls.Remove PalyaInfo.StartCelVonalNev.Label
        ' Nullázás.
        Set PalyaInfo.StartCelVonalNev.Label = Nothing
    End If

    ' Tömb törlése.
    ReDim PalyaInfo.PalyaVonalTomb(0 To 50) As VonalKoordinatak
    ' Tömb törlése.
    ReDim PalyaInfo.SzektorVonalTomb(0 To 3) As VonalKoordinatak
    ' Tömb törlése.
    ReDim PalyaInfo.KocsiVonalTomb(1 To 4) As VonalKoordinatak
    ' Tömb törlése.
    ReDim PalyaInfo.SzektorNevTomb(0 To 3) As NevKoordinatak

    ' Pálya vonalak számát nullára állítja.
    PalyaInfo.PalyaVonalakSzama = 0
    ' Szektor vonalak számát nullára állítja.
    PalyaInfo.SzektorVonalakSzama = 0
    ' Kocsi vonalak számát nullára állítja.
    PalyaInfo.KocsiVonalakSzama = 1
    ' Szektor nevek számát nullára állítja.
    PalyaInfo.SzektorNevekSzama = 0
End Sub
