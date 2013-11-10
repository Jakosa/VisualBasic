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

Private Const SaveDir = "Végeredmények"

Private Sub Class_Initialize()
    '
End Sub

Public Sub Save()
    If DirExists(SaveDir) Then
        MkDir SaveDir
    End If

    Dim Path As String
    Path = SaveDir & "/" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"

    Open Path For Output As #1

    Print #1, "# Mentett végeredmény."

    Dim tempkor As Byte, tempautok As Byte, ciklus As Integer, ciklus2 As Integer, i As Byte
    Dim NowTime As Date

    If Palya.GetKorokSzama > Config.Globalis_KorokSzama Then
        tempkor = Palya.GetKorokSzama - 1
    Else
        tempkor = Palya.GetKorokSzama
    End If

    tempautok = 0

    Do While True
        For ciklus = 3 To 1 Step -1
            For i = LBound(PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok) + tempautok To Palya.GetAutokSzama
                If PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin = "" And PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat Then
                    Exit For
                ElseIf PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat And tempautok <= Palya.GetAutokSzama Then
                    Print #1, ""
                    Print #1, "Sorrend: " & i & ". Autó: " & PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin

                    For ciklus2 = LBound(PalyaInfo.Autok) To Palya.GetAutokSzama
                        If PalyaInfo.Autok(ciklus2).GetColor = PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin Then
                            Print #1, "Legjobb idõ: " & PalyaInfo.Autok(ciklus2).GetLegjobbKorido & " másodperc"
                            Print #1, "Összes idõ idõ: " & PalyaInfo.Autok(ciklus2).GetOsszKorido & " másodperc"
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

                If tempautok = Palya.GetAutokSzama Then
                    Exit For
                End If
            Next i

            If tempautok = Palya.GetAutokSzama Then
                Exit For
            End If
        Next ciklus

        If tempautok = Palya.GetAutokSzama Then
            Exit Do
        End If

        If tempkor > Palya.GetKezdokorErteke Then
            tempkor = tempkor - 1
        Else
            Exit Do
        End If
    Loop

    Close #1
End Sub
