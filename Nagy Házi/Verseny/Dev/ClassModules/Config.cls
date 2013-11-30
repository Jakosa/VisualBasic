VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Konfig mappa nevét tárolja.
Private Const ConfigDir = "Config"
' A konfig fájl nevét tárolja.
Private Const ConfigFile = "Config.conf"
' A konfig fájl elérési útvonalát tárolja.
Private ConfigPath As String

' Globális változó. A neve takarja mira való.
Public Globalis_Nyomvonal As Boolean
' Globális változó. A neve takarja mira való.
Public Globalis_SzektorNevek As Boolean
' Globális változó. A neve takarja mira való.
Public Globalis_StartCelVonalNeve As Boolean
' Globális változó. A neve takarja mira való.
Public Globalis_KorokSzama As Byte
' Globális változó. A neve takarja mira való.
Public Globalis_PalyaNeve As String
' Globális változó. A neve takarja mira való.
Public Globalis_SzektorVonalak As Boolean
' Globális változó. A neve takarja mira való.
Public Globalis_StartCelVonal As Boolean
' Globális változó. A neve takarja mira való.
Public Globalis_TokeletesKorozes As Boolean
' Globális változó. A neve takarja mira való.
Public Autok_Elso_Nyomvonal As Boolean
' Globális változó. A neve takarja mira való.
Public Autok_Elso_TokeletesKorozes As Boolean
' Globális változó. A neve takarja mira való.
Public Autok_Masodik_Nyomvonal As Boolean
' Globális változó. A neve takarja mira való.
Public Autok_Masodik_TokeletesKorozes As Boolean
' Globális változó. A neve takarja mira való.
Public Autok_Harmadik_Nyomvonal As Boolean
' Globális változó. A neve takarja mira való.
Public Autok_Harmadik_TokeletesKorozes As Boolean
' Globális változó. A neve takarja mira való.
Public Autok_Negyedik_Nyomvonal As Boolean
' Globális változó. A neve takarja mira való.
Public Autok_Negyedik_TokeletesKorozes As Boolean

' Beállítjuk a form létrehozásakor az alap folyamatokat.
Private Sub Class_Initialize()
    ' Beállítja az elérési útvonalat.
    ConfigPath = ConfigDir & "/" & ConfigFile
End Sub

' Konfig fájl betöltése.
Public Sub LoadConfig()
    ' Megnézzi létezik-e a konfig mappája. Ha nem akkor létrehozza.
    If DirExists(ConfigDir) Then
        ' Konfig mappa létrehozása.
        MkDir ConfigDir
    End If

    ' Megnézzi létezik-e a konfig fájl.
    If FileExists(ConfigPath) Then
    ' Ha új konfig fájl lenne létrehozva akkor is elfut ez a rész hogy betöltödjenek a beállítások.
UjKonfigBetoltese:
        ' Konfig adatok alapértelmezésre állítása.
        SetDefault
        ' Fájl megnyítása olvasásra.
        Open ConfigPath For Input As #1

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
                    If InStr(Sor, "Globalis_Nyomvonal") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            Globalis_Nyomvonal = CBool(Parse(Sor))
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Globalis_Nyomvonal = AGlobalis_Nyomvonal
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Globalis_SzektorNevek") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            Globalis_SzektorNevek = CBool(Parse(Sor))
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Globalis_SzektorNevek = AGlobalis_SzektorNevek
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Globalis_StartCelVonalNeve") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            Globalis_StartCelVonalNeve = CBool(Parse(Sor))
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Globalis_StartCelVonalNeve = AGlobalis_StartCelVonalNeve
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Globalis_KorokSzama") > 0 Then
                        ' Megvizsgálja hogy "byte"-e az adat.
                        If IsByte(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            Globalis_KorokSzama = CByte(Parse(Sor))
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Globalis_KorokSzama = AGlobalis_KorokSzama
                        End If

                        ' Ha nagyobb a körök száma mint a maximális vagy a minimálisnál kisebb akkor az alapértelmezetett tölti be.
                        If Globalis_KorokSzama > 20 Or Globalis_KorokSzama < 2 Then
                            ' Alapértelmezett beállítás betöltése.
                            Globalis_KorokSzama = AGlobalis_KorokSzama
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Globalis_PalyaNeve") > 0 Then
                        ' Beállítás betöltése.
                        Globalis_PalyaNeve = CStr(Parse(Sor))
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Globalis_SzektorVonalak") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            Globalis_SzektorVonalak = CBool(Parse(Sor))
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Globalis_SzektorVonalak = AGlobalis_SzektorVonalak
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Globalis_SCV") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            Globalis_StartCelVonal = CBool(Parse(Sor))
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Globalis_StartCelVonal = AGlobalis_StartCelVonal
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Globalis_TokeletesKorozes") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            ' Globalis_TokeletesKorozes = CBool(Parse(Sor)) ' Kikapcsolva!!!
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Globalis_TokeletesKorozes = AGlobalis_TokeletesKorozes
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Autok_Elso_Nyomvonal") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            Autok_Elso_Nyomvonal = CBool(Parse(Sor))
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Autok_Elso_Nyomvonal = AAutok_Elso_Nyomvonal
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Autok_Elso_TokeletesKorozes") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            'Autok_Elso_TokeletesKorozes = CBool(Parse(Sor))  ' Kikapcsolva!!!
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Autok_Elso_TokeletesKorozes = AAutok_Elso_TokeletesKorozes
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Autok_Masodik_Nyomvonal") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            Autok_Masodik_Nyomvonal = CBool(Parse(Sor))
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Autok_Masodik_Nyomvonal = AAutok_Masodik_Nyomvonal
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Autok_Masodik_TokeletesKorozes") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            ' Autok_Masodik_TokeletesKorozes = CBool(Parse(Sor)) ' Kikapcsolva!!!
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Autok_Masodik_TokeletesKorozes = AAutok_Masodik_TokeletesKorozes
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Autok_Harmadik_Nyomvonal") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            Autok_Harmadik_Nyomvonal = CBool(Parse(Sor))
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Autok_Harmadik_Nyomvonal = AAutok_Harmadik_Nyomvonal
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Autok_Harmadik_TokeletesKorozes") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            ' Autok_Harmadik_TokeletesKorozes = CBool(Parse(Sor)) ' Kikapcsolva!!!
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Autok_Harmadik_TokeletesKorozes = AAutok_Harmadik_TokeletesKorozes
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Autok_Negyedik_Nyomvonal") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            Autok_Negyedik_Nyomvonal = CBool(Parse(Sor))
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Autok_Negyedik_Nyomvonal = AAutok_Negyedik_Nyomvonal
                        End If
                    End If

                    ' Akkor fut le ha létezik az adott beállítás.
                    If InStr(Sor, "Autok_Negyedik_TokeletesKorozes") > 0 Then
                        ' Megvizsgálja hogy "boolean"-e az adat.
                        If IsBoolean(Parse(Sor)) Then
                            ' Beállítás betöltése.
                            ' Autok_Negyedik_TokeletesKorozes = CBool(Parse(Sor)) ' Kikapcsolva!!!
                        Else
                            ' Alapértelmezett beállítás betöltése.
                            Autok_Negyedik_TokeletesKorozes = AAutok_Negyedik_TokeletesKorozes
                        End If
                    End If
                End If
            End If
        Loop

        ' Fájl lezására.
        Close #1
    Else
        ' Új konfig fájl létrehozása alapértelmezett beállításokkal.
        CreateConfig True

        ' Konfig beállítások betöltése.
        GoTo UjKonfigBetoltese
    End If
End Sub

' Konfig fájl létrehozása.
' Az "IsDefault" változó tárolja hogy alapértelmezet konfig fájlt kell-e létrehozni vagy a konfig fájlt kell módosítva létrehozni.
Private Sub CreateConfig(ByVal IsDefault As Boolean)
    Dim TempGlobalis_Nyomvonal As Boolean
    ' Ideiglenes változó.
    Dim TempGlobalis_SzektorNevek As Boolean
    ' Ideiglenes változó.
    Dim TempGlobalis_StartCelVonalNeve As Boolean
    ' Ideiglenes változó.
    Dim TempGlobalis_KorokSzama As Byte
    ' Ideiglenes változó.
    Dim TempGlobalis_PalyaNeve As String
    ' Ideiglenes változó.
    Dim TempGlobalis_SzektorVonalak As Boolean
    ' Ideiglenes változó.
    Dim TempGlobalis_StartCelVonal As Boolean
    ' Ideiglenes változó.
    Dim TempGlobalis_TokeletesKorozes As Boolean
    ' Ideiglenes változó.
    Dim TempAutok_Elso_Nyomvonal As Boolean
    ' Ideiglenes változó.
    Dim TempAutok_Elso_TokeletesKorozes As Boolean
    ' Ideiglenes változó.
    Dim TempAutok_Masodik_Nyomvonal As Boolean
    ' Ideiglenes változó.
    Dim TempAutok_Masodik_TokeletesKorozes As Boolean
    ' Ideiglenes változó.
    Dim TempAutok_Harmadik_Nyomvonal As Boolean
    ' Ideiglenes változó.
    Dim TempAutok_Harmadik_TokeletesKorozes As Boolean
    ' Ideiglenes változó.
    Dim TempAutok_Negyedik_Nyomvonal As Boolean
    ' Ideiglenes változó.
    Dim TempAutok_Negyedik_TokeletesKorozes As Boolean

    ' Akkor fut le ha az alapértelmezett beállításokat kell beállítani.
    If IsDefault Then
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempGlobalis_Nyomvonal = AGlobalis_Nyomvonal
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempGlobalis_SzektorNevek = AGlobalis_SzektorNevek
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempGlobalis_StartCelVonalNeve = AGlobalis_StartCelVonalNeve
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempGlobalis_KorokSzama = AGlobalis_KorokSzama
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempGlobalis_PalyaNeve = AGlobalis_PalyaNeve
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempGlobalis_SzektorVonalak = AGlobalis_SzektorVonalak
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempGlobalis_StartCelVonal = AGlobalis_StartCelVonal
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempGlobalis_TokeletesKorozes = AGlobalis_TokeletesKorozes
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempAutok_Elso_Nyomvonal = AAutok_Elso_Nyomvonal
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempAutok_Elso_TokeletesKorozes = AAutok_Elso_TokeletesKorozes
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempAutok_Masodik_Nyomvonal = AAutok_Masodik_Nyomvonal
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempAutok_Masodik_TokeletesKorozes = AAutok_Masodik_TokeletesKorozes
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempAutok_Harmadik_Nyomvonal = AAutok_Harmadik_Nyomvonal
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempAutok_Harmadik_TokeletesKorozes = AAutok_Harmadik_TokeletesKorozes
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempAutok_Negyedik_Nyomvonal = AAutok_Negyedik_Nyomvonal
        ' Alapértelmezett beállítás másolása az ideiglenes változóba.
        TempAutok_Negyedik_TokeletesKorozes = AAutok_Negyedik_TokeletesKorozes
    Else
        ' Beállítás másolása az ideiglenes változóba.
        TempGlobalis_Nyomvonal = Globalis_Nyomvonal
        ' Beállítás másolása az ideiglenes változóba.
        TempGlobalis_SzektorNevek = Globalis_SzektorNevek
        ' Beállítás másolása az ideiglenes változóba.
        TempGlobalis_StartCelVonalNeve = Globalis_StartCelVonalNeve
        ' Beállítás másolása az ideiglenes változóba.
        TempGlobalis_KorokSzama = Globalis_KorokSzama
        ' Beállítás másolása az ideiglenes változóba.
        TempGlobalis_PalyaNeve = Globalis_PalyaNeve
        ' Beállítás másolása az ideiglenes változóba.
        TempGlobalis_SzektorVonalak = Globalis_SzektorVonalak
        ' Beállítás másolása az ideiglenes változóba.
        TempGlobalis_StartCelVonal = Globalis_StartCelVonal
        ' Beállítás másolása az ideiglenes változóba.
        TempGlobalis_TokeletesKorozes = Globalis_TokeletesKorozes
        ' Beállítás másolása az ideiglenes változóba.
        TempAutok_Elso_Nyomvonal = Autok_Elso_Nyomvonal
        ' Beállítás másolása az ideiglenes változóba.
        TempAutok_Elso_TokeletesKorozes = Autok_Elso_TokeletesKorozes
        ' Beállítás másolása az ideiglenes változóba.
        TempAutok_Masodik_Nyomvonal = Autok_Masodik_Nyomvonal
        ' Beállítás másolása az ideiglenes változóba.
        TempAutok_Masodik_TokeletesKorozes = Autok_Masodik_TokeletesKorozes
        ' Beállítás másolása az ideiglenes változóba.
        TempAutok_Harmadik_Nyomvonal = Autok_Harmadik_Nyomvonal
        ' Beállítás másolása az ideiglenes változóba.
        TempAutok_Harmadik_TokeletesKorozes = Autok_Harmadik_TokeletesKorozes
        ' Beállítás másolása az ideiglenes változóba.
        TempAutok_Negyedik_Nyomvonal = Autok_Negyedik_Nyomvonal
        ' Beállítás másolása az ideiglenes változóba.
        TempAutok_Negyedik_TokeletesKorozes = Autok_Negyedik_TokeletesKorozes
    End If

    ' Fájl megnyítása írásra.
    Open ConfigPath For Output As #1

    ' Adatok kiírása a fájlba.
    Print #1, "# Játékhoz tartozó konfig fájl. Figyelem! Csak hozzáértõek piszkáljanak bele mert hibát is okozhatnak az átgondolatlan cselekvéssel!"
    Print #1, ""
    Print #1, "# Globálisan itt kapcsolható be vagy ki a nyomvonal."
    Print #1, "Globalis_Nyomvonal = " & CStr(TempGlobalis_Nyomvonal)
    Print #1, "# Globálisan itt kapcsolható be vagy ki a szektorok neve."
    Print #1, "Globalis_SzektorNevek = " & CStr(TempGlobalis_SzektorNevek)
    Print #1, "# Globálisan itt kapcsolható be vagy ki a start és célvonal neve."
    Print #1, "Globalis_StartCelVonalNeve = " & CStr(TempGlobalis_StartCelVonalNeve)
    Print #1, "# Itt állítható be hány körös legyen az adott verseny. Maximum: 20, Minimum: 2"
    Print #1, "Globalis_KorokSzama = " & TempGlobalis_KorokSzama
    Print #1, "# Itt állítható be a pálya neve."
    Print #1, "Globalis_PalyaNeve = " & TempGlobalis_PalyaNeve
    Print #1, "# Globálisan itt kapcsolható be vagy ki a szektorok vonala."
    Print #1, "Globalis_SzektorVonalak = " & CStr(TempGlobalis_SzektorVonalak)
    Print #1, "# Globálisan itt kapcsolható be vagy ki a start és cél vonala."
    Print #1, "Globalis_SCV = " & CStr(TempGlobalis_StartCelVonal)
    Print #1, "# Globálisan itt kapcsolható be vagy ki a tökéletes körözés."
    Print #1, "Globalis_TokeletesKorozes = " & CStr(TempGlobalis_TokeletesKorozes)
    Print #1, ""
    Print #1, "# Elsõ autónál itt kapcsolható be vagy ki a nyomvonal."
    Print #1, "Autok_Elso_Nyomvonal = " & CStr(TempAutok_Elso_Nyomvonal)
    Print #1, "# Elsõ autónál itt kapcsolható be vagy ki a tökéletes körözés."
    Print #1, "Autok_Elso_TokeletesKorozes = " & CStr(TempAutok_Elso_TokeletesKorozes)
    Print #1, "# Második autónál itt kapcsolható be vagy ki a nyomvonal."
    Print #1, "Autok_Masodik_Nyomvonal = " & CStr(TempAutok_Masodik_Nyomvonal)
    Print #1, "# Második autónál itt kapcsolható be vagy ki a tökéletes körözés."
    Print #1, "Autok_Masodik_TokeletesKorozes = " & CStr(TempAutok_Masodik_TokeletesKorozes)
    Print #1, "# Harmadik autónál itt kapcsolható be vagy ki a nyomvonal."
    Print #1, "Autok_Harmadik_Nyomvonal = " & CStr(TempAutok_Harmadik_Nyomvonal)
    Print #1, "# Harmadik autónál itt kapcsolható be vagy ki a tökéletes körözés."
    Print #1, "Autok_Harmadik_TokeletesKorozes = " & CStr(TempAutok_Harmadik_TokeletesKorozes)
    Print #1, "# Negyedik autónál itt kapcsolható be vagy ki a nyomvonal."
    Print #1, "Autok_Negyedik_Nyomvonal = " & CStr(TempAutok_Negyedik_Nyomvonal)
    Print #1, "# Negyedik autónál itt kapcsolható be vagy ki a tökéletes körözés."
    Print #1, "Autok_Negyedik_TokeletesKorozes = " & CStr(TempAutok_Negyedik_TokeletesKorozes)

    ' Fájl lezására.
    Close #1
End Sub

' Beállítja a módosított adatokat a konfig fájlban.
Public Sub SetConfig()
    ' Konfig fájl törlése.
    DeleteConfig
    ' Új konfig fájl létrehozása nem alapértelmezett beállításokkal.
    CreateConfig False
End Sub

' Törli a konfig fájlt.
Public Sub DeleteConfig()
    ' Konfig fájl törlése.
    Kill ConfigPath
End Sub

' Alapértelmezett beállítások beállítása.
Private Sub SetDefault()
    ' Alapértelmezett beállítása.
    Globalis_Nyomvonal = AGlobalis_Nyomvonal
    ' Alapértelmezett beállítása.
    Globalis_SzektorNevek = AGlobalis_SzektorNevek
    ' Alapértelmezett beállítása.
    Globalis_StartCelVonalNeve = AGlobalis_StartCelVonalNeve
    ' Alapértelmezett beállítása.
    Globalis_KorokSzama = AGlobalis_KorokSzama
    ' Alapértelmezett beállítása.
    Globalis_PalyaNeve = AGlobalis_PalyaNeve
    ' Alapértelmezett beállítása.
    Globalis_SzektorVonalak = AGlobalis_SzektorVonalak
    ' Alapértelmezett beállítása.
    Globalis_StartCelVonal = AGlobalis_StartCelVonal
    ' Alapértelmezett beállítása.
    Globalis_TokeletesKorozes = AGlobalis_TokeletesKorozes
    ' Alapértelmezett beállítása.
    Autok_Elso_Nyomvonal = AAutok_Elso_Nyomvonal
    ' Alapértelmezett beállítása.
    Autok_Elso_TokeletesKorozes = AAutok_Elso_TokeletesKorozes
    ' Alapértelmezett beállítása.
    Autok_Masodik_Nyomvonal = AAutok_Masodik_Nyomvonal
    ' Alapértelmezett beállítása.
    Autok_Masodik_TokeletesKorozes = AAutok_Masodik_TokeletesKorozes
    ' Alapértelmezett beállítása.
    Autok_Harmadik_Nyomvonal = AAutok_Harmadik_Nyomvonal
    ' Alapértelmezett beállítása.
    Autok_Harmadik_TokeletesKorozes = AAutok_Harmadik_TokeletesKorozes
    ' Alapértelmezett beállítása.
    Autok_Negyedik_Nyomvonal = AAutok_Negyedik_Nyomvonal
    ' Alapértelmezett beállítása.
    Autok_Negyedik_TokeletesKorozes = AAutok_Negyedik_TokeletesKorozes
End Sub

' Szétvágja a betöltött sorokat. Majd visszadja az értékes adatokat.
' A "Sor" változó a konfig fájl egy sorát tartalmazza ami éppen kiolvasásra került.
Private Function Parse(ByVal Sor As String) As String
    ' Egyenlõség jel utáni részt adja vissza. A felesleges adatokat eltávolítja (pl: szóköz).
    Parse = LTrim(Mid$(Sor, InStr(1, Sor, "=") + 1))
End Function
