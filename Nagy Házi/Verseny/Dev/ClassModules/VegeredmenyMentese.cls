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
        If Not PalyaInfo.Autok(i).GetGameEnd Then
            ' Figyelmeztetés ablak megnyítása.
            WarningWindow "Végeredmény mentése: Hiba!", "Nem futott végig a játék. Így nincs lehetõsége elmenteni a végeredményt.", False
            ' Kilépés az eljárásból.
            Exit Sub
        End If
    Next i

    If DirExists(SaveDir) Then
        MkDir SaveDir
    End If

    Dim Path As String
    Path = SaveDir & "/" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"

    Open Path For Output As #1

    Print #1, "# Mentett végeredmény."

    Dim tempkor As Byte, tempautok As Byte, ciklus As Integer, ciklus2 As Integer
    Dim NowTime As Date

    If Palya.GetKorokSzama > Config.Globalis_KorokSzama Then
        tempkor = Palya.GetKorokSzama - 1
    Else
        tempkor = Palya.GetKorokSzama
    End If

    tempautok = 0

    Do While True
        For ciklus = 3 To 1 Step -1
            For i = LBound(PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok) + tempautok To PalyaInfo.AutokSzama
                If PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin = "" And PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat Then
                    ' Kilépés a ciklusból.
                    Exit For
                ElseIf PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat And tempautok <= PalyaInfo.AutokSzama Then
                    Print #1, ""
                    Print #1, "Sorrend: " & i & ". Autó: " & PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin

                    For ciklus2 = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
                        If PalyaInfo.Autok(ciklus2).GetColor = PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin Then
                            Print #1, "Legjobb idõ: " & PalyaInfo.Autok(ciklus2).GetLegjobbKorido & " másodperc"
                            Print #1, "Összes idõ: " & PalyaInfo.Autok(ciklus2).GetOsszKorido & " másodperc"
                            Print #1, "Összes út: " & PalyaInfo.Autok(ciklus2).GetOsszesUt & " m"

                            If tempautok = 0 Then
                                NowTime = PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido
                                Print #1, "Autók közötti különbség: 0"
                            Else
                                Print #1, "Autók közötti különbség: +" & Abs(DateDiff("s", PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido, NowTime)) & " másodperc"
                            End If
                        End If
                    Next ciklus2

                    tempautok = tempautok + 1
                End If

                If tempautok = PalyaInfo.AutokSzama Then
                    ' Kilépés a ciklusból.
                    Exit For
                End If
            Next i

            If tempautok = PalyaInfo.AutokSzama Then
                ' Kilépés a ciklusból.
                Exit For
            End If
        Next ciklus

        If tempautok = PalyaInfo.AutokSzama Then
            ' Kilépés a ciklusból.
            Exit Do
        End If

        If tempkor > Palya.GetKezdokorErteke Then
            tempkor = tempkor - 1
        Else
            ' Kilépés a ciklusból.
            Exit Do
        End If
    Loop

    ' Fájl lezására.
    Close #1
End Sub
