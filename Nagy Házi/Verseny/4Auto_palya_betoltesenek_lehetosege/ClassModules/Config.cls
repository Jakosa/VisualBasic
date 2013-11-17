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
    
    If isdefault Then
        TempGlobalis_Nyomvonal = AGlobalis_Nyomvonal
        TempGlobalis_SzektorNevek = AGlobalis_SzektorNevek
        TempGlobalis_StartCelVonalNeve = AGlobalis_StartCelVonalNeve
        TempGlobalis_KorokSzama = AGlobalis_KorokSzama
        TempGlobalis_PalyaNeve = AGlobalis_PalyaNeve
    Else
        TempGlobalis_Nyomvonal = Globalis_Nyomvonal
        TempGlobalis_SzektorNevek = Globalis_SzektorNevek
        TempGlobalis_StartCelVonalNeve = Globalis_StartCelVonalNeve
        TempGlobalis_KorokSzama = Globalis_KorokSzama
        TempGlobalis_PalyaNeve = Globalis_PalyaNeve
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
End Sub

Private Function Parse(sor As String) As String
    Parse = LTrim(Mid$(sor, InStr(1, sor, "=") + 1))
End Function
