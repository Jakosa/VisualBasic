VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fejléc
' Készítette: Jakosa Csaba Árpád
' Fejléc vége

Option Explicit

' Az autó elhelyezkedését frissíti.
Private WithEvents Timer0 As VB.Timer
Attribute Timer0.VB_VarHelpID = -1
' A kocsi négy vonalát tárolja.
Private KVonalTomb(1 To 4) As VB.Line
' Véletlen nagyságú szövegeket tárol.
Private VSzogTomb(1 To 30) As Single
' Tárolja az autó színét.
Private Color As ColorConstants
' Tárolja hogy be illetve ki van-e kapcsolva a nyomvonal.
Private Nyomvonal As Boolean
' Tárolja hogy megy-e az autó vagy sem.
Private Started As Boolean
' X0 koordináta.
Private x0 As Double
' Y0 koordináta.
Private y0 As Double
Private ex As Double
Private ey As Double
' Léptéket tárol. A megadott lépték után majd a kód egyes részei változásokat idéznek elõ.
Private i As Integer
' Ha "true" akkor tökéletesebben középen van az autó. "false" esetén viszont szerintem élethübb.
Private PontosabbKor As Boolean
' Kocsi hossza.
Private Const xhossz = 80
' Kocsi szélessége.
Private Const yhossz = 60
' Tárolja a körök számát.
Private KorokSzama As Byte
' Egy kör idejének hossza.
Private EgyKorido As Date
' Tárolja azt az idõt ameddig az autó meg lett állítva.
Private StopKorido As Date
' Induláskor a köridõ beállításában segítõ változó.
Private BKorido As Boolean
' Megállást jelzõ változó.
Private BStopKorido As Boolean
' Egy szektor idejének hossza.
Private EgySzektorido As Date
' Tárolja hogy az autó átment-e az egyik szetoron vagy sem.
Private ASzektorvonal As Boolean
' Tárolja a három legjobb szektoridõt.
Private LegjobbSzektorido(1 To 3) As Single
' Tárolja a legjobb köridõt.
Private LegjobbKorido As Single
' Tárolja melyik körben lett beállítva a legjobb köridõ.
Private LegjobbKoridoSzama As Byte
' Tárolja hányas számú autó.
Private AutoSzama As Byte
' Tárolja vége van-e a játéknak vagy sem.
Private GameEnd As Boolean
' Ha vége a játéknak akkor ez a változó tárolja még mennyit fog az autó menni a pályán.
Private UtolsoMeterek As Integer
' Tárolja egy kör hosszát.
Private EgyKorHossza As Integer
' Tárolja az összes megtett útat.
Private OsszesUt As Integer
' Segíti a startvonal elõli indulást.
Private IndulasStartVonalElol As Boolean

' Tárolja a körök egyes adatait.
Private Type Korok
    ' Tárolja a köridõt.
    Korido As Single
    ' Tárolja a három szektoridõt.
    Szektorido(1 To 3) As Single
End Type

' Tárolja a köröket tömbben.
Private KorokTomb() As Korok

' Publikus változók.

' Visszadja publikusan hogy az autó befejezte-e a játékot.
Public Property Get GetGameEnd() As Boolean
    ' Érték beállítása.
    GetGameEnd = GameEnd
End Property

' Visszadja publikusan egy kör hosszát.
Public Property Get GetEgyKorHossza() As Integer
    ' Érték beállítása.
    GetEgyKorHossza = EgyKorHossza
End Property

' Visszadja publikusan az összes megtett útat.
Public Property Get GetOsszesUt() As Integer
    ' Érték beállítása.
    GetOsszesUt = OsszesUt
End Property

' Visszadja publikusan a körök számát.
Public Property Get GetKorokSzama() As Byte
    ' Érték beállítása.
    GetKorokSzama = KorokSzama
End Property

' Visszadja publikus változóként a kör adott szektoridejét.
' Az "si" tartalmazza a szektor számát.
Public Property Get GetKorSzektoridok(ByVal si As Integer) As Single
    ' Érték beállítása.
    GetKorSzektoridok = KorokTomb(KorokSzama).Szektorido(si)
End Property

' Visszadja publikusan az összes köridõt.
Public Property Get GetOsszKorido() As Single
    ' Tárolja az összes köridõt.
    Dim osszido As Single
    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Integer

    For ciklus = Palya.GetKezdokorErteke To Config.Globalis_KorokSzama
        ' Összes idõ összeadása.
        osszido = osszido + KorokTomb(ciklus).Korido
    Next ciklus

    ' Érték beállítása.
    GetOsszKorido = osszido
End Property

' Visszadja publikusan a legjobb köridõt.
Public Property Get GetLegjobbKorido() As Single
    ' Érték beállítása.
    GetLegjobbKorido = LegjobbKorido
End Property

' Visszadja publikusan a legjobb köridõ körének számát.
Public Property Get GetLegjobbKoridoSzama() As Byte
    ' Érték beállítása.
    GetLegjobbKoridoSzama = LegjobbKoridoSzama
End Property

' Visszadja publikus változóként a legjobb szektoridõt.
' Az "si" tartalmazza a szektor számát.
Public Property Get GetLegjobbSzektoridok(ByVal si As Integer) As Single
    ' Érték beállítása.
    GetLegjobbSzektoridok = LegjobbSzektorido(si)
End Property

' Visszadja publikus változóként a beállított színt.
Public Property Get GetColor() As String
    ' Tárolja a színt.
    Dim Szin As String

    ' Választás szín alapján.
    Select Case Color
        Case ColorConstants.vbBlue
            ' Szín beállítása.
            Szin = "kék"
        Case ColorConstants.vbRed
            ' Szín beállítása.
            Szin = "piros"
        Case ColorConstants.vbGreen
            ' Szín beállítása.
            Szin = "zöld"
        Case ColorConstants.vbYellow
            ' Szín beállítása.
            Szin = "sárga"
        Case ColorConstants.vbBlack
            ' Szín beállítása.
            Szin = "fekete"
        Case Else
            ' Szín beállítása.
            Szin = "fekete"
    End Select

    ' Érték beállítása.
    GetColor = Szin
End Property

' Visszadja publikus változóként az X0 koordinátát.
Public Property Get GetX0() As Double
    ' Érték beállítása.
    GetX0 = x0
End Property

' Visszadja publikus változóként az Y0 koordinátát.
Public Property Get GetY0() As Double
    ' Érték beállítása.
    GetY0 = y0
End Property

' Beállítható vele az X0 koordináta.
' Az "X" tartalmazza a koordinátát.
Public Sub SetX0(ByVal X As Double)
    ' Érték beállítása.
    x0 = KorrigacioX(X)
End Sub

' Beállítható vele az Y0 koordináta.
' Az "Y" tartalmazza a koordinátát.
Public Sub SetY0(ByVal Y As Double)
    ' Érték beállítása.
    y0 = KorrigacioY(Y)
End Sub

' Beállítható vele az EX koordináta.
' Az "X" tartalmazza a koordinátát.
Public Sub SetEX(ByVal X As Double)
    ' Érték beállítása.
    ex = X
End Sub

' Beállítható vele az EY koordináta.
' Az "Y" tartalmazza a koordinátát.
Public Sub SetEY(ByVal Y As Double)
    ' Érték beállítása.
    ey = Y
End Sub

' Beállítható vele az autó vonalainak színe.
' A "Szin" tárolja milyen színû legyen.
Public Sub SetColor(ByVal Szin As String)
    ' Kiválasztja a megfelelõ színt.
    Select Case Szin
        Case "Kék", "kék"
            ' Szín beállítása.
            Color = ColorConstants.vbBlue
        Case "Piros", "piros"
            ' Szín beállítása.
            Color = ColorConstants.vbRed
        Case "Zöld", "zöld"
            ' Szín beállítása.
            Color = ColorConstants.vbGreen
        Case "Sárga", "sárga"
            ' Szín beállítása.
            Color = ColorConstants.vbYellow
        Case "Fekete", "fekete"
            ' Szín beállítása.
            Color = ColorConstants.vbBlack
        Case Else
            ' Szín beállítása.
            Color = ColorConstants.vbBlack
    End Select

    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Byte
    For ciklus = LBound(KVonalTomb) To UBound(KVonalTomb)
        ' Érték beállítása.
        KVonalTomb(ciklus).BorderColor = Color
    Next ciklus
End Sub

' A vonalak láthatóságát lehet beállítani vele.
' A "Visible" tárolja hogy látható vagy láthatatlan legyen-e a vonal.
Public Sub SetVisible(ByVal Visible As Boolean)
    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Byte
    For ciklus = LBound(KVonalTomb) To UBound(KVonalTomb)
        ' Érték beállítása.
        KVonalTomb(ciklus).Visible = Visible
    Next ciklus
End Sub

' Beállítható vele a kocsi vonalának vastagsága.
' A "Nagysag" tartalmazza a vonal méretét.
Public Sub SetBorderWidth(ByVal Nagysag As Byte)
    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Byte
    For ciklus = LBound(KVonalTomb) To UBound(KVonalTomb)
        ' Érték beállítása.
        KVonalTomb(ciklus).BorderWidth = Nagysag
    Next ciklus
End Sub

' Nyomvonal be illetve kikapcsololásának lehetõsége.
' A "Nyom" tartalmazza azt hogy be vagy ki legyen-e kapcsolva.
Public Sub SetNyomvonal(ByVal Nyom As Boolean)
    ' Érték beállítása.
    Nyomvonal = Nyom
End Sub

' Be illetve kikapcsolhatóvá teszi hogy egy pontosabb kört írjon le az autó.
' A "PKor" tartalmazza hogy be illetve ki legyen-e kapcsolva ez a lehetõség.
Public Sub SetPontosabbKor(ByVal PKor As Boolean)
    ' Érték beállítása.
    PontosabbKor = PKor
End Sub

' Publikus változók vége.

' Betölti az autó egyedi információit.
' A "Szam" tartalmazza hányas számú autóként fog indulni.
Public Sub Load(ByVal Szam As Byte)
    ' Autó számának elmentése.
    AutoSzama = Szam

    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Byte
    For ciklus = LBound(KVonalTomb) To UBound(KVonalTomb)
        ' Kocsi vonal létrehozása.
        Set KVonalTomb(ciklus) = Palya.Controls.Add("VB.Line", "KocsiVonal" & Szam & CStr(ciklus), Palya.VirtualisPalya)
    Next ciklus

    ' Láthatóság beállítása.
    SetVisible True

    ' Idõzítõ létrehozása.
    Set Timer0 = Palya.Controls.Add("VB.Timer", "Timer" & Szam & "1", Palya)
    ' Idõzítõ kikapcsolása.
    Timer0.Enabled = False

    Randomize
    ' Érték beállítása. 50-71 millisec (Véletlen szám.)
    Timer0.Interval = Int(50 + Rnd * 21)

    For ciklus = LBound(VSzogTomb) To UBound(VSzogTomb)
        Randomize
        ' Véletlen nagyságú szög létrehozása.
        VSzogTomb(ciklus) = CSng(Rnd * 0.1)
    Next ciklus
End Sub

' Autó elindítása.
Public Sub Start()
    ' Akkor fut le ha az autó már megy vagy már befejezte a játékot.
    If Started Or GameEnd Then
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Akkor fut le ha új köridõt kell beállítani tehát igaz az értéke.
    If BKorido Then
        ' Jelenlegi idõ beállítása.
        EgyKorido = Now
        ' Az "EgyKorido" idejének átmásolása.
        EgySzektorido = EgyKorido
        ' Érték hamisra állítása.
        BKorido = False
    End If

    ' Érték igazra állítása. Az autó indul.
    Started = True
    ' Idözítõ elindítása.
    Timer0.Enabled = True
End Sub

' Autó megállítása.
Public Sub Stop_Kocsi()
    ' Akkor fut le ha az autó még nem fut vagy vége az autó játékának.
    If Not Started Or GameEnd Then
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Autó megállítása.
    Started = False
    ' Jelenlegi idõ beállítása.
    StopKorido = Now
    ' Megállást jelzõ változó igazra állítása és az állást mérõ változók bekapcsolása.
    BStopKorido = True
    ' Idõzítõ leállítása.
    Timer0.Enabled = False
End Sub

' Beállítjuk a form létrehozásakor az alap folyamatokat.
Private Sub Class_Initialize()
    ' Alapértelmezett érték beállítása.
    ex = 0.6
    ' Alapértelmezett érték beállítása.
    ey = -1
    ' Alapértelmezett érték beállítása.
    x0 = KorrigacioX(1100)
    ' Alapértelmezett érték beállítása.
    y0 = KorrigacioY(5000)
    ' Köridõ mértésének engedélyezése.
    BKorido = True
    ' Autó játékának befejztét jelzõ változó hamisra állítása.
    GameEnd = False
    ' Autó megállását jelzõ változó hamisra állítása.
    BStopKorido = False
    ' Az elsõ vonaltol való indulást jelzõ változó hamisra állítása.
    IndulasStartVonalElol = False
    ' Kezdõkör beállítása.
    KorokSzama = Palya.GetKezdokorErteke()
    ' Nyomvonal bekapcsolása.
    SetNyomvonal True
    ' Pontosabb körözés kikapcsolása.
    SetPontosabbKor False
    ' Szektorvonalon való átmenetel hamisra állítása.
    ASzektorvonal = False
    ' Tömbben tárol körök törlése és újradimenzionálása.
    ReDim KorokTomb(Palya.GetKezdokorErteke() To Config.Globalis_KorokSzama) As Korok

    ' Kezdõérték beállítása.
    LegjobbKorido = KezdoSzektorido
    ' Kezdõérték beállítása.
    i = LBound(VSzogTomb)

    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Integer
    For ciklus = LBound(LegjobbSzektorido) To UBound(LegjobbSzektorido)
        ' Kezdõérték beállítása.
        LegjobbSzektorido(ciklus) = KezdoSzektorido
    Next ciklus
End Sub

' Class megszünésekor bizonyos dolgok megsemisítésre kerülnek.
Private Sub Class_Terminate()
    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Byte
    For ciklus = LBound(KVonalTomb) To UBound(KVonalTomb)
        ' Nullázás.
        Set KVonalTomb(ciklus) = Nothing
    Next ciklus

    ' Nullázás.
    Set Timer0 = Nothing
End Sub

' Törli az egyes információkat a kocsikról. Megsemmisíti õket.
Public Sub Dispose()
    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Byte
    For ciklus = LBound(KVonalTomb) To UBound(KVonalTomb)
        ' Törli a kocsi vonalát a pályáról.
        Palya.Controls.Remove KVonalTomb(ciklus)
    Next ciklus

    ' Törli az idõzítöt.
    Palya.Controls.Remove Timer0
End Sub

' Megjeleníti az autót.
Public Sub Show()
    ' Autó vonalainak mozgatása a következõ koordinátákra.
    NextMove
End Sub

' Korrigálja az X eltérést.
' A "PointX" tárolja hol található az X koordináta.
Private Function KorrigacioX(ByVal PointX As Single) As Single
    ' A virtuális pálya x tengely távolságának levonása.
    KorrigacioX = PointX - Palya.VirtualisPalya.Left
End Function

' Korrigálja az Y eltérést.
' A "PointY" tárolja hol található az Y koordináta.
Private Function KorrigacioY(ByVal PointY As Single) As Single
    ' A virtuális pálya y tengely távolságának levonása.
    KorrigacioY = PointY - Palya.VirtualisPalya.Top
End Function

' Két szektoridõ mérése értelemszerüen távolság alapján ha elérte.
' 0-tól 1-ig tartanak a szektorvonalak mert a célvonal egyben a harmadik szektor.
Private Sub Szektorok()
    ' Akkor fut le ha a körök száma már nagyobb mint ami be lett állítva a konfigban.
    If KorokSzama > Config.Globalis_KorokSzama Then
        ' Kilépés az eljárásból.
        Exit Sub
    End If

    ' Tárolja hogy új kör kezdõdik-e.
    Dim NextKor As Boolean
    ' "ii" segédváltozó a ciklushoz.
    Dim ii As Byte

    For ii = LBound(KorokTomb(KorokSzama).Szektorido) To UBound(KorokTomb(KorokSzama).Szektorido)
        ' Tárolja a szektorvonaltól való távolságot.
        Dim dist As Single
        ' "ciklus" segédváltozó a ciklushoz.
        Dim ciklus As Byte
        ' Tárolja a szektor számát.
        Dim Szam As Byte
        ' Kezdõérték beállítása a távolságnak.
        dist = KezdoTavolsagPontok

        For ciklus = LBound(PalyaInfo.SzektorVonalTomb) To PalyaInfo.SzektorVonalakSzama - 1
            ' Kocsivonal távolságának segédváltozója.
            Dim d As Single
            ' Távolság beállítása.
            d = Distance(x0, y0, KorrigacioX(PalyaInfo.SzektorVonalTomb(ciklus).X1), KorrigacioX(PalyaInfo.SzektorVonalTomb(ciklus).X2), KorrigacioY(PalyaInfo.SzektorVonalTomb(ciklus).Y1), KorrigacioY(PalyaInfo.SzektorVonalTomb(ciklus).Y2))

            ' A legkisebb távolságot keresi.
            If dist > d Then
                ' Érték beálíltása.
                dist = d
                ' Szektor számának beállítása.
                Szam = ciklus + 1
            End If
        Next ciklus

        ' Akkor fut le ha a távolság kisebb mint 100 és nem ment át az autó egy szektorvonalon se.
        If dist < 100 And Not ASzektorvonal Then
            ' Tárolni fogja a jelenlegi idõt.
            Dim NowTime As Date
            ' Jelenlegi idõ beállítása.
            NowTime = Now
            ' A szektorvonalon való átmenés igazra állítása.
            ASzektorvonal = True

            ' Akkor fut le ha nem a startvonal elöt van az autó.
            If Not IndulasStartVonalElol Then
                ' Most már a start vonal után lesz.
                IndulasStartVonalElol = True
                ' Kilépés az eljárásból.
                Exit Sub
            End If

            For ciklus = LBound(PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).Autok) To UBound(PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).Autok)
                ' Akkor fut le ha még nincs adat az adott szektorban.
                If Not PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).VanAdat Then
                    ' Érték igazra állítása.
                    PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).VanAdat = True
                End If

                ' Akkor fut le ha az autó színe még nincs beállítva a szektorban.
                If PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).Autok(ciklus).Szin = "" Then
                    ' Szín beállítása.
                    PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).Autok(ciklus).Szin = GetColor
                    ' Idõ beállítása.
                    PalyaInfo.SorrendTomb(KorokSzama).Szektor(Szam).Autok(ciklus).Ido = NowTime
                    ' Kilépés a ciklusból.
                    Exit For
                End If
            Next ciklus

            ' Akkor fut le ha a harmadik célvonalon ment át az autó. Az az a célvonalon.
            If Szam = 3 Then
                ' Új kör lehetõségének engedélyezése.
                NextKor = True
                ' Köridõ beállítása.
                KorokTomb(KorokSzama).Korido = DateDiff("s", EgyKorido, NowTime)

                ' Akkor fut le ha az autó körideje jobb mint a régi.
                If LegjobbKorido > KorokTomb(KorokSzama).Korido Then
                    ' Új legjobb köridõ beállítása.
                    LegjobbKorido = KorokTomb(KorokSzama).Korido
                    ' Legjobb köridõ kör számának beállítása.
                    LegjobbKoridoSzama = KorokSzama
                End If

                ' Teljes köridõ beállítása az utolsó szektornak.
                KorokTomb(KorokSzama).Szektorido(Szam) = KorokTomb(KorokSzama).Korido

                ' "j" segédváltozó a ciklushoz.
                Dim j As Byte
                For j = LBound(KorokTomb(KorokSzama).Szektorido) To UBound(KorokTomb(KorokSzama).Szektorido) - 1
                    ' Az utolsó szektoridõt úgy kapjuk meg hogy a másikakat kivonjuk belõle.
                    KorokTomb(KorokSzama).Szektorido(Szam) = KorokTomb(KorokSzama).Szektorido(Szam) - KorokTomb(KorokSzama).Szektorido(j)
                Next j
            Else
                ' Szektoridõ beállítása.
                KorokTomb(KorokSzama).Szektorido(Szam) = DateDiff("s", EgySzektorido, NowTime)
            End If

            ' Akkor fut le ha az autó szektorideje jobb mint a régi.
            If LegjobbSzektorido(Szam) > KorokTomb(KorokSzama).Szektorido(Szam) Then
                ' Új legjobb szektoridõ beállítása.
                LegjobbSzektorido(Szam) = KorokTomb(KorokSzama).Szektorido(Szam)
            End If

            ' Jelenlegi idõ beállítása.
            NowTime = Now

            ' Akkor fut le ha a harmadik célvonalon ment át az autó. Az az a célvonalon.
            If Szam = 3 Then
                'Jelenlegi idõ beállítása.
                EgyKorido = NowTime
            End If

            ' Jelenlegi idõ beállítása.
            EgySzektorido = NowTime
        ' Akkor fut le ha a távolság nagyobb mint 100 és átment az autó egy szektorvonalon.
        ElseIf dist > 100 And ASzektorvonal Then
            ' Szektorvonalon való átmenés hamisra állítása.
            ASzektorvonal = False
        End If
    Next ii

    ' Akkor fut le ha új kör jön.
    If NextKor Then
        ' Körök számának megnövelése eggyel.
        KorokSzama = KorokSzama + 1
    End If

    ' Akkor fut le ha a körök száma már nagyobb mint ami be lett állítva a konfigban.
    If KorokSzama > Config.Globalis_KorokSzama Then
        ' Autó játékát befejezetre állítja.
        GameEnd = True
    End If
End Sub

' Timer0 idõzitõ Timer eseménye.
Private Sub Timer0_Timer()
    ' Érték növelése egyyel.
    i = i + 1

    ' Akkor fut le ha a tömbb elemeinek számánál nagyobb az "i".
    If i > UBound(VSzogTomb) - LBound(VSzogTomb) + 1 Then
        ' "i" visszaállítása az alap értékre.
        i = LBound(VSzogTomb)

        Randomize
        ' Érték beállítása. 50-71 millisec (Véletlen szám.)
        Timer0.Interval = Int(50 + Rnd * 21)
    End If

    ' Akkor fut le ha az autó meg lett állítva.
    If BStopKorido Then
        ' Tárolni fogja a jelenlegi idõt.
        Dim NowTime As Date
        ' Jelenlegi idõ beállítása.
        NowTime = Now
        ' Idõk közötti különbség kiszámítása.
        EgyKorido = DateAdd("s", DateDiff("s", StopKorido, NowTime), EgyKorido)
        ' Idõk közötti különbség kiszámítása.
        EgySzektorido = DateAdd("s", DateDiff("s", StopKorido, NowTime), EgySzektorido)
        ' Érték hamisra állítása.
        BStopKorido = False
    End If

    ' Balra kanyarodás véletlen szöggel.
    NextCoordinate True, True

    ' Tárolja a baloldali volnal X koordinátáját.
    Dim xb As Single
    ' Tárolja a baloldali volnal Y koordinátáját.
    Dim yb As Single
    ' Tárolja a jobboldali volnal X koordinátáját.
    Dim xj As Single
    ' Tárolja a jobboldali volnal Y koordinátáját.
    Dim yj As Single

    ' Ideálisan középen van így a jármû.
    If PontosabbKor Then
        ' XB beállítása.
        xb = x0 + xhossz * ex - yhossz * ey
        ' YB beállítása.
        yb = y0 + xhossz * ey + yhossz * ex
        ' XJ beállítása.
        xj = x0 + xhossz * ex + yhossz * ey
        ' YJ beállítása.
        yj = y0 + xhossz * ey - yhossz * ex
    Else
        ' XB beállítása.
        xb = x0 + 400 * ex - 300 * ey
        ' YB beállítása.
        yb = y0 + 400 * ey + 300 * ex
        ' XJ beállítása.
        xj = x0 + 400 * ex + 300 * ey
        ' YJ beállítása.
        yj = y0 + 400 * ey - 300 * ex
    End If

    ' Baloldali kocsivonaltól mért távolságot tárolja.
    Dim distb As Single
    ' Jobboldali kocsivonaltól mért távolságot tárolja.
    Dim distj As Single
    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Integer
    ' Kezdõérték beállítása a távolságnak.
    distb = KezdoTavolsagPontok
    ' Kezdõérték beállítása a távolságnak.
    distj = KezdoTavolsagPontok

    For ciklus = LBound(PalyaInfo.PalyaVonalTomb) To PalyaInfo.PalyaVonalakSzama - 1
        ' Baloldali kocsivonal távolságának segédváltozója.
        Dim db As Single
        ' Jobboldali kocsivonal távolságának segédváltozója.
        Dim dj As Single
        ' Távolság beállítása.
        db = Distance(xb, yb, KorrigacioX(PalyaInfo.PalyaVonalTomb(ciklus).X1), KorrigacioX(PalyaInfo.PalyaVonalTomb(ciklus).X2), KorrigacioY(PalyaInfo.PalyaVonalTomb(ciklus).Y1), KorrigacioY(PalyaInfo.PalyaVonalTomb(ciklus).Y2))
        ' Távolság beállítása.
        dj = Distance(xj, yj, KorrigacioX(PalyaInfo.PalyaVonalTomb(ciklus).X1), KorrigacioX(PalyaInfo.PalyaVonalTomb(ciklus).X2), KorrigacioY(PalyaInfo.PalyaVonalTomb(ciklus).Y1), KorrigacioY(PalyaInfo.PalyaVonalTomb(ciklus).Y2))

        ' A legkisebb távolságot keresi.
        If distb > db Then
             ' Érték beállítása.
            distb = db
        End If

        ' A legkisebb távolságot keresi.
        If distj > dj Then
            ' Érték beállítása.
            distj = dj
        End If
    Next ciklus

    ' Ideálisan középen van így a jármû.
    If PontosabbKor Then
        ' Akkor fut le ha a baloldali és jobboldali kocsivonal különbségétõl mért távolság kisebb mint 200 de nagyobb mint nulla.
        If distb - distj > 0 And distb - distj < 200 Then
            ' Jobbra fordulás.
            NextCoordinate False, False
            ' Autó vonalainak mozgatása a következõ koordinátákra.
            NextMove
        ' Akkor fut le ha a jobboldali és baloldali kocsivonal különbségétõl mért távolság kisebb mint 200 de nagyobb mint nulla.
        ElseIf distj - distb > 0 And distj - distb < 200 Then
            ' Balra fordulás.
            NextCoordinate True, False
            ' Autó vonalainak mozgatása a következõ koordinátákra.
            NextMove
        End If
    Else
        ' Akkor fut le ha a baloldali kocsivonaltol mért távolság kisebb mint 200.
        If distb < 200 Then
            ' Balra fordulás.
            NextCoordinate True, False
            ' Autó vonalainak mozgatása a következõ koordinátákra.
            NextMove
        ' Akkor fut le ha a baloldali kocsivonaltol mért távolság nagyobb mint 200.
        ElseIf distb > 200 Then
            ' Jobbra fordulás.
            NextCoordinate False, False
            ' Autó vonalainak mozgatása a következõ koordinátákra.
            NextMove
        Else
            ' X0 beállítása.
            x0 = x0 - 55 * ex
            ' Y0 beállítása.
            y0 = y0 - 55 * ey
        End If
    End If

    ' Autó vonalainak mozgatása a következõ koordinátákra.
    NextMove

    ' X0 beállítása.
    x0 = x0 + 50 * ex
    ' Y0 beállítása.
    y0 = y0 + 50 * ey
    ' Autó vonalainak mozgatása a következõ koordinátákra.
    NextMove

    ' Szektorok elemzése.
    Szektorok

    ' Nyomvonal bekapcsolása ha igaz az érték.
    If Nyomvonal Then
        ' Nyomvonal kirajzolása.
        Palya.VirtualisPalya.Circle (x0, y0), 10, Color
    End If

    ' Akkor fut le ha az elsõ körben vagyunk.
    If KorokSzama = Palya.GetKezdokorErteke And IndulasStartVonalElol Then
        ' Egy kör hosszának mérése.
        EgyKorHossza = EgyKorHossza + PalyaHosszanakLepteke
    End If

    ' Összes megtett út számítása.
    OsszesUt = OsszesUt + PalyaHosszanakLepteke

    ' Akkor fut le ha a kocsi célbaért már.
    If GameEnd Then
        ' Utolsó méterek számolása léptékkel.
        UtolsoMeterek = UtolsoMeterek + 1

        ' Akkor fut le ha az utolsó méterek száma elérte a maximumot.
        If UtolsoMeterek = 20 Then
            ' Idõzítõ kikapcsolása.
            Timer0.Enabled = False
        End If
    End If
End Sub

' Elmozdítja a koordinátákat a következõ pontra.
' A "BalraKanyarodas" azt tárolja hogy az autó balra vagy jobbra forduljon-e. Ha balra kell fordulni akkor igaz ha jobbra akkor hamis.
' A "VSzog" igazra állításával a véletlen szögeket használja fel az eljárás a tömbbõl.
Private Sub NextCoordinate(ByVal BalraKanyarodas As Boolean, ByVal VSzog As Boolean)
    ' Tárolja a balra illetve a jobbra kanyarodás mértékét.
    Dim bj As Double

    ' Akkor fut le ha igaz a változó. Ebben az esetben balra kanyarodik az autó.
    If BalraKanyarodas Then
        ' Balra kanyarodás mértéke.
        bj = -0.15
    Else
        ' Jobbra kanyarodás mértéke.
        bj = 0.15
    End If

    ' Ha igaz akkor a tömbben tárolt véletlen szöggel fordul el.
    If VSzog Then
        ' X komponens.
        ex = Cos(VSzogTomb(i)) * ex - Sin(VSzogTomb(i)) * ey
        ' Y komponens.
        ey = Cos(VSzogTomb(i)) * ey + Sin(VSzogTomb(i)) * ex
    Else
        ' X komponens.
        ex = Cos(bj) * ex - Sin(bj) * ey
        ' Y komponens.
        ey = Cos(bj) * ey + Sin(bj) * ex
    End If

    ' X komponens mértékének korrigációja.
    ex = ex / Sqr(ex * ex + ey * ey)
    ' Y komponens mértékének korrigációja.
    ey = ey / Sqr(ex * ex + ey * ey)
End Sub

' Mozgatás a következõ pontokba.
Private Sub NextMove()
    ' Vonal X1 pontjának legenerálása.
    KVonalTomb(1).X1 = x0 + xhossz * ex - yhossz * ey
    ' Vonal Y1 pontjának legenerálása.
    KVonalTomb(1).Y1 = y0 + xhossz * ey + yhossz * ex
    ' Vonal X2 pontjának legenerálása.
    KVonalTomb(1).X2 = x0 - xhossz * ex - yhossz * ey
    ' Vonal Y2 pontjának legenerálása.
    KVonalTomb(1).Y2 = y0 - xhossz * ey + yhossz * ex

    ' Vonal X1 pontjának legenerálása.
    KVonalTomb(2).X1 = x0 + xhossz * ex + yhossz * ey
    ' Vonal Y1 pontjának legenerálása.
    KVonalTomb(2).Y1 = y0 + xhossz * ey - yhossz * ex
    ' Vonal X2 pontjának legenerálása.
    KVonalTomb(2).X2 = x0 - xhossz * ex + yhossz * ey
    ' Vonal Y2 pontjának legenerálása.
    KVonalTomb(2).Y2 = y0 - xhossz * ey - yhossz * ex

    ' Vonal X1 pontjának legenerálása.
    KVonalTomb(3).X1 = KVonalTomb(1).X1
    ' Vonal X2 pontjának legenerálása.
    KVonalTomb(3).X2 = KVonalTomb(2).X1
    ' Vonal Y1 pontjának legenerálása.
    KVonalTomb(3).Y1 = KVonalTomb(1).Y1
    ' Vonal Y2 pontjának legenerálása.
    KVonalTomb(3).Y2 = KVonalTomb(2).Y1

    ' Vonal X1 pontjának legenerálása.
    KVonalTomb(4).X1 = KVonalTomb(1).X2
    ' Vonal X2 pontjának legenerálása.
    KVonalTomb(4).X2 = KVonalTomb(2).X2
    ' Vonal Y1 pontjának legenerálása.
    KVonalTomb(4).Y1 = KVonalTomb(1).Y2
    ' Vonal Y2 pontjának legenerálása.
    KVonalTomb(4).Y2 = KVonalTomb(2).Y2
End Sub
