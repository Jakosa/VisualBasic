VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VegeredmenyMentese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Végeredmény mappájának neve.
Private Const SaveDir = "Végeredmények"

' A végeredmény elmentésére szolgál.
Public Sub Save()
    ' "i" segédváltozó a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Akkor fut le ha valamelyik autónak még nem ért véget a játék.
        If Not PalyaInfo.Autok(i).GetGameEnd Then
            ' Figyelmeztetés ablak megnyítása.
            WarningWindow "Végeredmény mentése: Hiba!", "Nem futott végig a játék. Így nincs lehetõsége elmenteni a végeredményt.", False
            ' Kilépés az eljárásból.
            Exit Sub
        End If
    Next i

    ' Megnézi létezik-e a könyvtár.
    If DirExists(SaveDir) Then
        ' Könyvtár létrehozása.
        MkDir SaveDir
    End If

    ' Tárolja az elérési útvonalat.
    Dim Path As String
    ' Elérési útvonal beállítása.
    Path = SaveDir & "/" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"

    ' Fájl megnyítása írásra.
    Open Path For Output As #1

    ' Adat kiírása fájlba.
    Print #1, "# Mentett végeredmény."

    ' Ideiglenes köröket tárol.
    Dim tempkor As Byte
    ' Ideiglenes autók számát tárolja.
    Dim tempautok As Byte
    ' "ciklus" segédváltozó a ciklushoz.
    Dim ciklus As Integer
    ' "ciklus2" segédváltozó a ciklushoz.
    Dim ciklus2 As Integer
    ' Tárolja a szektor idõt.
    Dim NowTime As Date

    ' Ha a Palya.GetKorokSzama nagyobb mint a maximális körök száma akkor fut le.
    If Palya.GetKorokSzama > Config.Globalis_KorokSzama Then
        ' Érték beállítása. Azért -1 mert a változó a játék végén +1-el nagyobbra lett megnövelve.
        tempkor = Palya.GetKorokSzama - 1
    Else
        ' Érték beállítása.
        tempkor = Palya.GetKorokSzama
    End If

    ' Nullázás.
    tempautok = 0

    ' Végtelenségig futó ciklus
    Do While True
        For ciklus = 3 To 1 Step -1
            For i = LBound(PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok) + tempautok To PalyaInfo.AutokSzama
                ' Akkor fut le ha nincs szin beállítva (nincs autó) és a van adat is.
                If PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin = "" And PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat Then
                    ' Kilépés a ciklusból.
                    Exit For
                ' Akkor fut le ha van adat és az ideiglenes autók száma kisebb vagy engyenlõ az AutokSzama-val.
                ElseIf PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat And tempautok <= PalyaInfo.AutokSzama Then
                    ' Adat kiírása fájlba.
                    Print #1, ""
                    ' Adat kiírása fájlba.
                    Print #1, "Sorrend: " & i & ". Autó: " & PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin

                    For ciklus2 = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
                        ' Akkor fut le ha a kocsi szine egyenlõ a szektorhoz tartózó kocsi szinével.
                        If PalyaInfo.Autok(ciklus2).GetColor = PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin Then
                            ' Adat kiírása fájlba.
                            Print #1, "Legjobb idõ: " & PalyaInfo.Autok(ciklus2).GetLegjobbKorido & " másodperc"
                            ' Adat kiírása fájlba.
                            Print #1, "Összes idõ: " & PalyaInfo.Autok(ciklus2).GetOsszKorido & " másodperc"
                            ' Adat kiírása fájlba.
                            Print #1, "Összes út: " & PalyaInfo.Autok(ciklus2).GetOsszesUt & " m"

                            ' Akkor fut le ha az ideiglenes autók száma nem nulla.
                            If tempautok = 0 Then
                                ' Menti a szektor idejét.
                                NowTime = PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido
                                ' Adat kiírása fájlba.
                                Print #1, "Autók közötti különbség: 0"
                            Else
                                ' Adat kiírása fájlba.
                                Print #1, "Autók közötti különbség: +" & Abs(DateDiff("s", PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido, NowTime)) & " másodperc"
                            End If
                        End If
                    Next ciklus2

                    ' Megnöveljük 1-el az ideiglenes autók számát.
                    tempautok = tempautok + 1
                End If

                ' Akkor fut le ha az ideiglenes autók száma egyenlõ az AutokSzama-val.
                If tempautok = PalyaInfo.AutokSzama Then
                    ' Kilépés a ciklusból.
                    Exit For
                End If
            Next i

            ' Akkor fut le ha az ideiglenes autók száma egyenlõ az AutokSzama-val.
            If tempautok = PalyaInfo.AutokSzama Then
                ' Kilépés a ciklusból.
                Exit For
            End If
        Next ciklus

        ' Akkor fut le ha az ideiglenes autók száma egyenlõ az AutokSzama-val.
        If tempautok = PalyaInfo.AutokSzama Then
            ' Kilépés a ciklusból.
            Exit Do
        End If

        ' Akkor fut le ha az ideiglenes körök száma nagyobb mind a kezdõkör értéke.
        If tempkor > Palya.GetKezdokorErteke Then
            ' Az ideiglenes körök számát csökkentjük eggyel.
            tempkor = tempkor - 1
        Else
            ' Kilépés a ciklusból.
            Exit Do
        End If
    Loop

    ' Fájl lezására.
    Close #1
End Sub
