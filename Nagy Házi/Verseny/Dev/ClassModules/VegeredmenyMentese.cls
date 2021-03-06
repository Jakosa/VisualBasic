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
' Fejl�c
' K�sz�tette: Salamon Valter
' Fejl�c v�ge

Option Explicit

' V�geredm�ny mapp�j�nak neve.
Private Const SaveDir = "V�geredm�nyek"

' A v�geredm�ny elment�s�re szolg�l.
Public Sub Save()
    ' "i" seg�dv�ltoz� a ciklushoz.
    Dim i As Byte
    For i = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
        ' Akkor fut le ha valamelyik aut�nak m�g nem �rt v�get a j�t�k.
        If Not PalyaInfo.Autok(i).GetGameEnd Then
            ' Figyelmeztet�s ablak megny�t�sa.
            WarningWindow "V�geredm�ny ment�se: Hiba!", "Nem futott v�gig a j�t�k. �gy nincs lehet�s�ge elmenteni a v�geredm�nyt.", False
            ' Kil�p�s az elj�r�sb�l.
            Exit Sub
        End If
    Next i

    ' Megn�zi l�tezik-e a k�nyvt�r.
    If DirExists(SaveDir) Then
        ' K�nyvt�r l�trehoz�sa.
        MkDir SaveDir
    End If

    ' T�rolja az el�r�si �tvonalat.
    Dim Path As String
    ' El�r�si �tvonal be�ll�t�sa.
    Path = SaveDir & "/" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"

    ' F�jl megny�t�sa �r�sra.
    Open Path For Output As #1

    ' Adat ki�r�sa f�jlba.
    Print #1, "# Mentett v�geredm�ny."

    ' Ideiglenes k�r�ket t�rol.
    Dim tempkor As Byte
    ' Ideiglenes aut�k sz�m�t t�rolja.
    Dim tempautok As Byte
    ' "ciklus" seg�dv�ltoz� a ciklushoz.
    Dim ciklus As Integer
    ' "ciklus2" seg�dv�ltoz� a ciklushoz.
    Dim ciklus2 As Integer
    ' T�rolja a szektor id�t.
    Dim NowTime As Date

    ' Ha a Palya.GetKorokSzama nagyobb mint a maxim�lis k�r�k sz�ma akkor fut le.
    If Palya.GetKorokSzama > Config.Globalis_KorokSzama Then
        ' �rt�k be�ll�t�sa. Az�rt -1 mert a v�ltoz� a j�t�k v�g�n +1-el nagyobbra lett megn�velve.
        tempkor = Palya.GetKorokSzama - 1
    Else
        ' �rt�k be�ll�t�sa.
        tempkor = Palya.GetKorokSzama
    End If

    ' Null�z�s.
    tempautok = 0

    ' V�gtelens�gig fut� ciklus
    Do While True
        For ciklus = 3 To 1 Step -1
            For i = LBound(PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok) + tempautok To PalyaInfo.AutokSzama
                ' Akkor fut le ha nincs szin be�ll�tva (nincs aut�) �s a van adat is.
                If PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin = "" And PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat Then
                    ' Kil�p�s a ciklusb�l.
                    Exit For
                ' Akkor fut le ha van adat �s az ideiglenes aut�k sz�ma kisebb vagy engyenl� az AutokSzama-val.
                ElseIf PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).VanAdat And tempautok <= PalyaInfo.AutokSzama Then
                    ' Adat ki�r�sa f�jlba.
                    Print #1, ""
                    ' Adat ki�r�sa f�jlba.
                    Print #1, "Sorrend: " & i & ". Aut�: " & PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin

                    For ciklus2 = LBound(PalyaInfo.Autok) To PalyaInfo.AutokSzama
                        ' Akkor fut le ha a kocsi szine egyenl� a szektorhoz tart�z� kocsi szin�vel.
                        If PalyaInfo.Autok(ciklus2).GetColor = PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Szin Then
                            ' Adat ki�r�sa f�jlba.
                            Print #1, "Legjobb id�: " & PalyaInfo.Autok(ciklus2).GetLegjobbKorido & " m�sodperc"
                            ' Adat ki�r�sa f�jlba.
                            Print #1, "�sszes id�: " & PalyaInfo.Autok(ciklus2).GetOsszKorido & " m�sodperc"
                            ' Adat ki�r�sa f�jlba.
                            Print #1, "�sszes �t: " & PalyaInfo.Autok(ciklus2).GetOsszesUt & " m"

                            ' Akkor fut le ha az ideiglenes aut�k sz�ma nem nulla.
                            If tempautok = 0 Then
                                ' Menti a szektor idej�t.
                                NowTime = PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido
                                ' Adat ki�r�sa f�jlba.
                                Print #1, "Aut�k k�z�tti k�l�nbs�g: 0"
                            Else
                                ' Adat ki�r�sa f�jlba.
                                Print #1, "Aut�k k�z�tti k�l�nbs�g: +" & Abs(DateDiff("s", PalyaInfo.SorrendTomb(tempkor).Szektor(ciklus).Autok(i).Ido, NowTime)) & " m�sodperc"
                            End If
                        End If
                    Next ciklus2

                    ' Megn�velj�k 1-el az ideiglenes aut�k sz�m�t.
                    tempautok = tempautok + 1
                End If

                ' Akkor fut le ha az ideiglenes aut�k sz�ma egyenl� az AutokSzama-val.
                If tempautok = PalyaInfo.AutokSzama Then
                    ' Kil�p�s a ciklusb�l.
                    Exit For
                End If
            Next i

            ' Akkor fut le ha az ideiglenes aut�k sz�ma egyenl� az AutokSzama-val.
            If tempautok = PalyaInfo.AutokSzama Then
                ' Kil�p�s a ciklusb�l.
                Exit For
            End If
        Next ciklus

        ' Akkor fut le ha az ideiglenes aut�k sz�ma egyenl� az AutokSzama-val.
        If tempautok = PalyaInfo.AutokSzama Then
            ' Kil�p�s a ciklusb�l.
            Exit Do
        End If

        ' Akkor fut le ha az ideiglenes k�r�k sz�ma nagyobb mind a kezd�k�r �rt�ke.
        If tempkor > Palya.GetKezdokorErteke Then
            ' Az ideiglenes k�r�k sz�m�t cs�kkentj�k eggyel.
            tempkor = tempkor - 1
        Else
            ' Kil�p�s a ciklusb�l.
            Exit Do
        End If
    Loop

    ' F�jl lez�s�ra.
    Close #1
End Sub
