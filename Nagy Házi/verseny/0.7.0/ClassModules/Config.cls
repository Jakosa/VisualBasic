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

Private Const ConfigDir = "Config"
Private Const ConfigFile = "Config.conf"
Private ConfigPath As String

Public Globalis_Nyomvonal As Boolean
Public Globalis_SzektorNevek As Boolean
Public Globalis_StartCelVonalNeve As Boolean
Public Globalis_KorokSzama As Byte
Public Globalis_PalyaNeve As String
Public Globalis_SzektorVonalak As Boolean
Public Globalis_StartCelVonal As Boolean
Public Globalis_TokeletesKorozes As Boolean
Public Autok_Elso_Nyomvonal As Boolean
Public Autok_Elso_TokeletesKorozes As Boolean
Public Autok_Masodik_Nyomvonal As Boolean
Public Autok_Masodik_TokeletesKorozes As Boolean
Public Autok_Harmadik_Nyomvonal As Boolean
Public Autok_Harmadik_TokeletesKorozes As Boolean
Public Autok_Negyedik_Nyomvonal As Boolean
Public Autok_Negyedik_TokeletesKorozes As Boolean

Private Sub Class_Initialize()
    ConfigPath = ConfigDir & "/" & ConfigFile
End Sub

Public Sub LoadConfig()
    If DirExists(ConfigDir) Then
        MkDir ConfigDir
    End If

    If FileExists(ConfigPath) Then
UjKonfigBetoltese:
        SetDefault
        Open ConfigPath For Input As #1

        Do Until EOF(1)
            Dim sor As String
            Line Input #1, sor

            If Len(sor) > 0 Then
                If Not Mid$(sor, 1, 1) = "#" Then
                    If InStr(sor, "Globalis_Nyomvonal") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Globalis_Nyomvonal = CBool(Parse(sor))
                        Else
                            Globalis_Nyomvonal = AGlobalis_Nyomvonal
                        End If
                    End If

                    If InStr(sor, "Globalis_SzektorNevek") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Globalis_SzektorNevek = CBool(Parse(sor))
                        Else
                            Globalis_SzektorNevek = AGlobalis_SzektorNevek
                        End If
                    End If

                    If InStr(sor, "Globalis_StartCelVonalNeve") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Globalis_StartCelVonalNeve = CBool(Parse(sor))
                        Else
                            Globalis_StartCelVonalNeve = AGlobalis_StartCelVonalNeve
                        End If
                    End If

                    If InStr(sor, "Globalis_KorokSzama") > 0 Then
                        If IsByte(Parse(sor)) Then
                            Globalis_KorokSzama = CByte(Parse(sor))
                        Else
                            Globalis_KorokSzama = AGlobalis_KorokSzama
                        End If

                        If Globalis_KorokSzama > 20 Then
                            Globalis_KorokSzama = AGlobalis_KorokSzama
                        End If
                    End If

                    If InStr(sor, "Globalis_PalyaNeve") > 0 Then
                        Globalis_PalyaNeve = CStr(Parse(sor))
                    End If

                    If InStr(sor, "Globalis_SzektorVonalak") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Globalis_SzektorVonalak = CBool(Parse(sor))
                        Else
                            Globalis_SzektorVonalak = AGlobalis_SzektorVonalak
                        End If
                    End If

                    If InStr(sor, "Globalis_SCV") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Globalis_StartCelVonal = CBool(Parse(sor))
                        Else
                            Globalis_StartCelVonal = AGlobalis_StartCelVonal
                        End If
                    End If

                    If InStr(sor, "Globalis_TokeletesKorozes") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Globalis_TokeletesKorozes = CBool(Parse(sor))
                        Else
                            Globalis_TokeletesKorozes = AGlobalis_TokeletesKorozes
                        End If
                    End If

                    If InStr(sor, "Autok_Elso_Nyomvonal") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Autok_Elso_Nyomvonal = CBool(Parse(sor))
                        Else
                            Autok_Elso_Nyomvonal = AAutok_Elso_Nyomvonal
                        End If
                    End If

                    If InStr(sor, "Autok_Elso_TokeletesKorozes") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Autok_Elso_TokeletesKorozes = CBool(Parse(sor))
                        Else
                            Autok_Elso_TokeletesKorozes = AAutok_Elso_TokeletesKorozes
                        End If
                    End If

                    If InStr(sor, "Autok_Masodik_Nyomvonal") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Autok_Masodik_Nyomvonal = CBool(Parse(sor))
                        Else
                            Autok_Masodik_Nyomvonal = AAutok_Masodik_Nyomvonal
                        End If
                    End If

                    If InStr(sor, "Autok_Masodik_TokeletesKorozes") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Autok_Masodik_TokeletesKorozes = CBool(Parse(sor))
                        Else
                            Autok_Masodik_TokeletesKorozes = AAutok_Masodik_TokeletesKorozes
                        End If
                    End If

                    If InStr(sor, "Autok_Harmadik_Nyomvonal") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Autok_Harmadik_Nyomvonal = CBool(Parse(sor))
                        Else
                            Autok_Harmadik_Nyomvonal = AAutok_Harmadik_Nyomvonal
                        End If
                    End If

                    If InStr(sor, "Autok_Harmadik_TokeletesKorozes") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Autok_Harmadik_TokeletesKorozes = CBool(Parse(sor))
                        Else
                            Autok_Harmadik_TokeletesKorozes = AAutok_Harmadik_TokeletesKorozes
                        End If
                    End If

                    If InStr(sor, "Autok_Negyedik_Nyomvonal") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Autok_Negyedik_Nyomvonal = CBool(Parse(sor))
                        Else
                            Autok_Negyedik_Nyomvonal = AAutok_Negyedik_Nyomvonal
                        End If
                    End If

                    If InStr(sor, "Autok_Negyedik_TokeletesKorozes") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Autok_Negyedik_TokeletesKorozes = CBool(Parse(sor))
                        Else
                            Autok_Negyedik_TokeletesKorozes = AAutok_Negyedik_TokeletesKorozes
                        End If
                    End If
                End If
            End If
        Loop
        Close #1
    Else
        CreateConfig True

        GoTo UjKonfigBetoltese
    End If
End Sub

Private Sub CreateConfig(isdefault As Boolean)
    Dim TempGlobalis_Nyomvonal As Boolean
    Dim TempGlobalis_SzektorNevek As Boolean
    Dim TempGlobalis_StartCelVonalNeve As Boolean
    Dim TempGlobalis_KorokSzama As Byte
    Dim TempGlobalis_PalyaNeve As String
    Dim TempGlobalis_SzektorVonalak As Boolean
    Dim TempGlobalis_StartCelVonal As Boolean
    Dim TempGlobalis_TokeletesKorozes As Boolean
    Dim TempAutok_Elso_Nyomvonal As Boolean
    Dim TempAutok_Elso_TokeletesKorozes As Boolean
    Dim TempAutok_Masodik_Nyomvonal As Boolean
    Dim TempAutok_Masodik_TokeletesKorozes As Boolean
    Dim TempAutok_Harmadik_Nyomvonal As Boolean
    Dim TempAutok_Harmadik_TokeletesKorozes As Boolean
    Dim TempAutok_Negyedik_Nyomvonal As Boolean
    Dim TempAutok_Negyedik_TokeletesKorozes As Boolean
    
    If isdefault Then
        TempGlobalis_Nyomvonal = AGlobalis_Nyomvonal
        TempGlobalis_SzektorNevek = AGlobalis_SzektorNevek
        TempGlobalis_StartCelVonalNeve = AGlobalis_StartCelVonalNeve
        TempGlobalis_KorokSzama = AGlobalis_KorokSzama
        TempGlobalis_PalyaNeve = AGlobalis_PalyaNeve
        TempGlobalis_SzektorVonalak = AGlobalis_SzektorVonalak
        TempGlobalis_StartCelVonal = AGlobalis_StartCelVonal
        TempGlobalis_TokeletesKorozes = AGlobalis_TokeletesKorozes
        TempAutok_Elso_Nyomvonal = AAutok_Elso_Nyomvonal
        TempAutok_Elso_TokeletesKorozes = AAutok_Elso_TokeletesKorozes
        TempAutok_Masodik_Nyomvonal = AAutok_Masodik_Nyomvonal
        TempAutok_Masodik_TokeletesKorozes = AAutok_Masodik_TokeletesKorozes
        TempAutok_Harmadik_Nyomvonal = AAutok_Harmadik_Nyomvonal
        TempAutok_Harmadik_TokeletesKorozes = AAutok_Harmadik_TokeletesKorozes
        TempAutok_Negyedik_Nyomvonal = AAutok_Negyedik_Nyomvonal
        TempAutok_Negyedik_TokeletesKorozes = AAutok_Negyedik_TokeletesKorozes
    Else
        TempGlobalis_Nyomvonal = Globalis_Nyomvonal
        TempGlobalis_SzektorNevek = Globalis_SzektorNevek
        TempGlobalis_StartCelVonalNeve = Globalis_StartCelVonalNeve
        TempGlobalis_KorokSzama = Globalis_KorokSzama
        TempGlobalis_PalyaNeve = Globalis_PalyaNeve
        TempGlobalis_SzektorVonalak = Globalis_SzektorVonalak
        TempGlobalis_StartCelVonal = Globalis_StartCelVonal
        TempGlobalis_TokeletesKorozes = Globalis_TokeletesKorozes
        TempAutok_Elso_Nyomvonal = Autok_Elso_Nyomvonal
        TempAutok_Elso_TokeletesKorozes = Autok_Elso_TokeletesKorozes
        TempAutok_Masodik_Nyomvonal = Autok_Masodik_Nyomvonal
        TempAutok_Masodik_TokeletesKorozes = Autok_Masodik_TokeletesKorozes
        TempAutok_Harmadik_Nyomvonal = Autok_Harmadik_Nyomvonal
        TempAutok_Harmadik_TokeletesKorozes = Autok_Harmadik_TokeletesKorozes
        TempAutok_Negyedik_Nyomvonal = Autok_Negyedik_Nyomvonal
        TempAutok_Negyedik_TokeletesKorozes = Autok_Negyedik_TokeletesKorozes
    End If

    Open ConfigPath For Output As #1

    Print #1, "# Játékhoz tartozó konfig fájl. Figyelem! Csak hozzáértõek piszkáljanak bele mert hibát is okozhatnak az átgondolatlan cselekvéssel!"
    Print #1, ""
    Print #1, "# Globálisan itt kapcsolható be vagy ki a nyomvonal."
    Print #1, "Globalis_Nyomvonal = " & CStr(TempGlobalis_Nyomvonal)
    Print #1, "# Globálisan itt kapcsolható be vagy ki a szektorok neve."
    Print #1, "Globalis_SzektorNevek = " & CStr(TempGlobalis_SzektorNevek)
    Print #1, "# Globálisan itt kapcsolható be vagy ki a start és célvonal neve."
    Print #1, "Globalis_StartCelVonalNeve = " & CStr(TempGlobalis_StartCelVonalNeve)
    Print #1, "# Itt állítható be hány körös legyen az adott verseny. Maximum 20."
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
    Close #1
End Sub

Public Sub SetConfig()
    DeleteConfig
    CreateConfig False
End Sub

Public Sub DeleteConfig()
    Kill ConfigPath
End Sub

Private Sub SetDefault()
    Globalis_Nyomvonal = AGlobalis_Nyomvonal
    Globalis_SzektorNevek = AGlobalis_SzektorNevek
    Globalis_StartCelVonalNeve = AGlobalis_StartCelVonalNeve
    Globalis_KorokSzama = AGlobalis_KorokSzama
    Globalis_PalyaNeve = AGlobalis_PalyaNeve
    Globalis_SzektorVonalak = AGlobalis_SzektorVonalak
    Globalis_StartCelVonal = AGlobalis_StartCelVonal
    Globalis_TokeletesKorozes = AGlobalis_TokeletesKorozes
    Autok_Elso_Nyomvonal = AAutok_Elso_Nyomvonal
    Autok_Elso_TokeletesKorozes = AAutok_Elso_TokeletesKorozes
    Autok_Masodik_Nyomvonal = AAutok_Masodik_Nyomvonal
    Autok_Masodik_TokeletesKorozes = AAutok_Masodik_TokeletesKorozes
    Autok_Harmadik_Nyomvonal = AAutok_Harmadik_Nyomvonal
    Autok_Harmadik_TokeletesKorozes = AAutok_Harmadik_TokeletesKorozes
    Autok_Negyedik_Nyomvonal = AAutok_Negyedik_Nyomvonal
    Autok_Negyedik_TokeletesKorozes = AAutok_Negyedik_TokeletesKorozes
End Sub

Private Function Parse(sor As String) As String
    Parse = LTrim(Mid$(sor, InStr(1, sor, "=") + 1))
End Function
