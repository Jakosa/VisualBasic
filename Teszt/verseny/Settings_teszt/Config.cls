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
Public Globalis_KorokSzama As Byte

Private Sub Class_Initialize()
    ConfigPath = ConfigDir & "/" & ConfigFile
End Sub

Public Sub LoadConfig()
    If DirExists(ConfigDir) Then
        MkDir ConfigDir
    End If

    If FileExists(ConfigPath) Then
UjKonfigBetoltese:
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
                    Else
                        Globalis_Nyomvonal = AGlobalis_Nyomvonal
                    End If

                    If InStr(sor, "Globalis_SzektorNevek") > 0 Then
                        If IsBoolean(Parse(sor)) Then
                            Globalis_SzektorNevek = CBool(Parse(sor))
                        Else
                            Globalis_SzektorNevek = AGlobalis_SzektorNevek
                        End If
                    Else
                        Globalis_SzektorNevek = AGlobalis_SzektorNevek
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
                    Else
                        Globalis_KorokSzama = AGlobalis_KorokSzama
                    End If
                End If
            End If
        Loop

        Close #1
    Else
        CreateConfig

        GoTo UjKonfigBetoltese
    End If
End Sub

Private Sub CreateConfig()
    Open ConfigPath For Output As #1

    Print #1, "# Játékhoz tartozó konfig fájl. Figyelem! Csak hozzáértõek piszkáljanak bele mert hibát is okozhatnak az átgondolatlan cselekvéssel!"
    Print #1, ""
    Print #1, "# Globálisan itt kapcsolható be vagy ki a nyomvonal."
    Print #1, "Globalis_Nyomvonal = " & CStr(AGlobalis_Nyomvonal)
    Print #1, "# Globálisan itt kapcsolható be vagy ki a szektorok neve."
    Print #1, "Globalis_SzektorNevek = " & CStr(AGlobalis_SzektorNevek)
    Print #1, "# Itt állítható be hány körös legyen az adott verseny. Maximum 20."
    Print #1, "Globalis_KorokSzama = " & AGlobalis_KorokSzama
    Close #1
End Sub

Public Sub SetConfig()
    DeleteConfig
    Open ConfigPath For Output As #1

    Print #1, "# Játékhoz tartozó konfig fájl. Figyelem! Csak hozzáértõek piszkáljanak bele mert hibát is okozhatnak az átgondolatlan cselekvéssel!"
    Print #1, ""
    Print #1, "# Globálisan itt kapcsolható be vagy ki a nyomvonal."
    Print #1, "Globalis_Nyomvonal = " & CStr(Globalis_Nyomvonal)
    Print #1, "# Globálisan itt kapcsolható be vagy ki a szektorok neve."
    Print #1, "Globalis_SzektorNevek = " & CStr(Globalis_SzektorNevek)
    Print #1, "# Itt állítható be hány körös legyen az adott verseny. Maximum 20."
    Print #1, "Globalis_KorokSzama = " & Globalis_KorokSzama
    Close #1
End Sub

Public Sub DeleteConfig()
    Kill ConfigPath
End Sub

Private Function Parse(sor As String) As String
    Parse = LTrim(Mid$(sor, InStr(1, sor, "=") + 1))
End Function
