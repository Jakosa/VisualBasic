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

Private MapPath As String

Private Sub Class_Initialize()
    MapPath = MapDir & "/" & Config.Globalis_PalyaNeve
End Sub

Public Sub LoadMap(ByVal MapName As String)
    Clean

    If DirExists(MapDir) Then
        MkDir MapDir
    End If

    MapPath = MapDir & "/" & MapName

    If FileExists(MapPath) Then
UjBetoltese:
        Open MapPath For Input As #1

        Do Until EOF(1)
            Dim Sor As String
            Line Input #1, Sor

            If Len(Sor) > 0 Then
                If Not Mid$(Sor, 1, 1) = "#" Then
                    If InStr(Sor, "PalyaVonal") > 0 Then
                        If PalyaInfo.PalyaVonalakSzama >= UBound(PalyaInfo.PalyaVonalTomb) Then
                            ReDim Preserve PalyaInfo.PalyaVonalTomb(0 To UBound(PalyaInfo.PalyaVonalTomb) + 50) As VonalKoordinatak
                        End If

                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama) = ParseLine(Parse(Sor))
                        Set PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal = Palya.Controls.Add("VB.Line", "PalyaVonal" & CStr(PalyaInfo.PalyaVonalakSzama), Palya)
                        'PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.BorderColor = &H80000008
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.X1 = PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).X1
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.X2 = PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).X2
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.Y1 = PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Y1
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.Y2 = PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Y2
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.BorderWidth = 2
                        PalyaInfo.PalyaVonalTomb(PalyaInfo.PalyaVonalakSzama).Vonal.Visible = True
                        PalyaInfo.PalyaVonalakSzama = PalyaInfo.PalyaVonalakSzama + 1
                    End If

                    If InStr(Sor, "SzektorVonal") > 0 Then
                        If PalyaInfo.SzektorVonalakSzama >= UBound(PalyaInfo.SzektorVonalTomb) Then
                            ReDim Preserve PalyaInfo.SzektorVonalTomb(0 To UBound(PalyaInfo.SzektorVonalTomb) + 5) As VonalKoordinatak
                        End If

                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama) = ParseLine(Parse(Sor))
                        Set PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal = Palya.Controls.Add("VB.Line", "SzektorVonal" & CStr(PalyaInfo.SzektorVonalakSzama), Palya)
                        'PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.BorderColor = &H80000008
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.X1 = PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).X1
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.X2 = PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).X2
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.Y1 = PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Y1
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.Y2 = PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Y2
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.BorderWidth = 1
                        PalyaInfo.SzektorVonalTomb(PalyaInfo.SzektorVonalakSzama).Vonal.Visible = True
                        PalyaInfo.SzektorVonalakSzama = PalyaInfo.SzektorVonalakSzama + 1
                    End If

                    If InStr(Sor, "SzektorNev") > 0 Then
                        If PalyaInfo.SzektorNevekSzama >= UBound(PalyaInfo.SzektorNevTomb) Then
                            ReDim Preserve PalyaInfo.SzektorNevTomb(0 To UBound(PalyaInfo.SzektorNevTomb) + 5) As NevKoordinatak
                        End If

                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama) = ParseLabel(Parse(Sor))
                        Set PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label = Palya.Controls.Add("VB.Label", "SzektorNev" & CStr(PalyaInfo.SzektorNevekSzama), Palya)
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.BackColor = Palya.BackColor
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.AutoSize = True
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.Caption = "Szektor " & CStr(PalyaInfo.SzektorNevekSzama + 1)
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.Left = PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Left
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.Top = PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Top
                        PalyaInfo.SzektorNevTomb(PalyaInfo.SzektorNevekSzama).Label.Visible = True
                        PalyaInfo.SzektorNevekSzama = PalyaInfo.SzektorNevekSzama + 1
                    End If

                    If InStr(Sor, "StartCelVonalNev") > 0 Then
                        PalyaInfo.StartCelVonalNev = ParseLabel(Parse(Sor))
                        Set PalyaInfo.StartCelVonalNev.Label = Palya.Controls.Add("VB.Label", "StartCelVonalNev", Palya)
                        PalyaInfo.StartCelVonalNev.Label.BackColor = Palya.BackColor
                        PalyaInfo.StartCelVonalNev.Label.AutoSize = True
                        PalyaInfo.StartCelVonalNev.Label.Caption = "Start / Cél"
                        PalyaInfo.StartCelVonalNev.Label.Left = PalyaInfo.StartCelVonalNev.Left
                        PalyaInfo.StartCelVonalNev.Label.Top = PalyaInfo.StartCelVonalNev.Top
                        PalyaInfo.StartCelVonalNev.Label.Visible = True
                    End If

                    If InStr(Sor, "KorokSzama") > 0 Then
                        If IsByte(Parse(Sor)) Then
                            PalyaInfo.KorokSzama = CByte(Parse(Sor))
                        Else
                            PalyaInfo.KorokSzama = AGlobalis_KorokSzama
                        End If

                        If PalyaInfo.KorokSzama > 20 Then
                            PalyaInfo.KorokSzama = AGlobalis_KorokSzama
                        End If
                    End If

                    If InStr(Sor, "KocsiVonal") > 0 Then
                        If PalyaInfo.KocsiVonalakSzama >= UBound(PalyaInfo.SzektorVonalTomb) Then
                            ReDim Preserve PalyaInfo.KocsiVonalTomb(1 To UBound(PalyaInfo.KocsiVonalTomb) + 4) As VonalKoordinatak
                        End If

                        PalyaInfo.KocsiVonalTomb(PalyaInfo.KocsiVonalakSzama) = ParseLine(Parse(Sor))
                        PalyaInfo.KocsiVonalakSzama = PalyaInfo.KocsiVonalakSzama + 1
                    End If
                End If
            End If
        Loop

        Close #1
    Else
        MapPath = MapDir & "/" & AGlobalis_PalyaNeve

        If Not FileExists(MapPath) Then
            CreateMap
        End If

        GoTo UjBetoltese
    End If
End Sub

Public Sub DeleteDefaultMap()
    MapPath = MapDir & "/" & AGlobalis_PalyaNeve

    If FileExists(MapPath) Then
        Kill MapPath
    End If
End Sub

Private Sub CreateMap()
    Open MapPath For Output As #1

    Print #1, "# Alapértelmezett map."
    Print #1, ""
    Print #1, "# Pályához tartozó körök száma."
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
    Close #1
End Sub

Private Function Parse(ByVal Sor As String) As String
    Parse = LTrim(Mid$(Sor, InStr(1, Sor, ":") + 1))
End Function

Private Function ParseLine(ByVal Sor As String) As VonalKoordinatak
    Dim vk As VonalKoordinatak
    Dim v() As String

    v = Split(Sor, ";")

    If UBound(v) - LBound(v) + 1 < 4 Then
        Exit Function
    End If

    If IsInteger(v(0)) Then
        vk.X1 = v(0)
    Else
        vk.X1 = 0
    End If

    If IsInteger(v(1)) Then
        vk.X2 = v(1)
    Else
        vk.X2 = 0
    End If

    If IsInteger(v(2)) Then
        vk.Y1 = v(2)
    Else
        vk.Y1 = 0
    End If

    If IsInteger(v(3)) Then
        vk.Y2 = v(3)
    Else
        vk.Y2 = 0
    End If

    ParseLine = vk
End Function

Private Function ParseLabel(ByVal Sor As String) As NevKoordinatak
    Dim nk As NevKoordinatak
    Dim v() As String

    v = Split(Sor, ";")

    If UBound(v) - LBound(v) + 1 < 2 Then
        Exit Function
    End If

    If IsInteger(v(0)) Then
        nk.Left = v(0)
    Else
        nk.Left = 0
    End If

    If IsInteger(v(1)) Then
        nk.Top = v(1)
    Else
        nk.Top = 0
    End If

    ParseLabel = nk
End Function

Private Sub Clean()
    Dim i As Integer

    If PalyaInfo.PalyaVonalakSzama > 0 Then
        For i = LBound(PalyaInfo.PalyaVonalTomb) To PalyaInfo.PalyaVonalakSzama - 1
            Palya.Controls.Remove PalyaInfo.PalyaVonalTomb(i).Vonal
            Set PalyaInfo.PalyaVonalTomb(i).Vonal = Nothing
        Next i
    End If

    If PalyaInfo.SzektorVonalakSzama > 0 Then
        For i = LBound(PalyaInfo.SzektorVonalTomb) To PalyaInfo.SzektorVonalakSzama - 1
            Palya.Controls.Remove PalyaInfo.SzektorVonalTomb(i).Vonal
            Set PalyaInfo.SzektorVonalTomb(i).Vonal = Nothing
        Next i
    End If

    If PalyaInfo.SzektorNevekSzama > 0 Then
        For i = LBound(PalyaInfo.SzektorNevTomb) To PalyaInfo.SzektorNevekSzama - 1
            Palya.Controls.Remove PalyaInfo.SzektorNevTomb(i).Label
            Set PalyaInfo.SzektorNevTomb(i).Label = Nothing
        Next i
    End If

    If Not PalyaInfo.StartCelVonalNev.Label Is Nothing Then
        Palya.Controls.Remove PalyaInfo.StartCelVonalNev.Label
        Set PalyaInfo.StartCelVonalNev.Label = Nothing
    End If

    ReDim PalyaInfo.PalyaVonalTomb(0 To 50) As VonalKoordinatak
    ReDim PalyaInfo.SzektorVonalTomb(0 To 3) As VonalKoordinatak
    ReDim PalyaInfo.KocsiVonalTomb(1 To 4) As VonalKoordinatak
    ReDim PalyaInfo.SzektorNevTomb(0 To 3) As NevKoordinatak

    PalyaInfo.PalyaVonalakSzama = 0
    PalyaInfo.SzektorVonalakSzama = 0
    PalyaInfo.KocsiVonalakSzama = 1
    PalyaInfo.SzektorNevekSzama = 0
End Sub
